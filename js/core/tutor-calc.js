/*
 * 導生座談：檢核與餐費補助計算（不含畫面，可在瀏覽器與 Node 測試共用）
 *
 * 規則（使用者確認 2026-11-09）：
 *  - 狀態：rawdata 當月有座談紀錄 → V；有簽到（美德耐各類 + 青果森）→ 院內。
 *    「院外」由使用者手動填寫，程式不判斷。
 *  - 簽到但 rawdata 當月無紀錄 → 院內，紅底白字標示（臨培無 rawdata，不標示）。
 *  - 餐費：臨培 / PGY / R 每人 180 元，UGY（實習生）每人 360 元；青果森依學員類別分 R+PGY 與 UGY。
 *  - 重複申請：同一卡號當月簽到超過一次。保留第一筆（美德耐優先，其次青果森），其餘列為重複。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TutorCalc = factory();
})(typeof self !== 'undefined' ? self : this, () => {
  const RATES = { 臨培: 180, UGY: 360, PGY: 180, R: 180 };
  const CAT_ORDER = ['R', 'PGY', 'UGY', '臨培'];
  const CAT_LABEL = { R: '住院醫師', PGY: 'PGY', UGY: '實習生', 臨培: '臨培（醫事人員）' };
  /** 簽到列的識別碼（重複申請選擇用） */
  const signKey = (s) => `${s.file}|${s.page}|${s.line}`;

  const str = (v) => (v === null || v === undefined ? '' : String(v)).normalize('NFKC').trim();

  /** 卡號：大寫；'1234D'（4 碼數字 + 檢查碼）→ '1234'；純數字補到 4 碼 */
  function normCard(v) {
    let s = str(v).toUpperCase().replace(/\s/g, '');
    if (/^\d{4}[A-Z]$/.test(s)) s = s.slice(0, 4);
    if (/^\d{1,3}$/.test(s)) s = s.padStart(4, '0');
    return s;
  }

  function normDate(v) {
    if (v instanceof Date && !isNaN(v)) {
      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
    }
    const m = str(v).match(/^(\d{3,4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (!m) return '';
    let y = Number(m[1]); if (y < 1911) y += 1911;
    return `${y}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  }

  /** 卡號前綴推測類別（名冊與 rawdata 都找不到時才用） */
  function guessCat(card) {
    if (/^U/.test(card)) return 'PGY';
    if (/^A/.test(card)) return 'R';
    if (/^7\d{3}$/.test(card)) return 'UGY';
    return '';
  }

  // ---------- rawdata ----------
  /** 依表頭判斷 rawdata 類別 */
  function detectRawType(headers) {
    const h = headers.map(str);
    if (!h.includes('學員卡號') || !h.includes('座談日期')) return null;
    if (h.includes('學校')) return 'UGY';
    if (h.includes('單位')) return 'R';
    return 'PGY';
  }

  /** rawdata 物件陣列 → 全部紀錄（不限月份） */
  function parseRaw(rows, cat) {
    return rows.map((r) => ({
      cat,
      card: normCard(r['學員卡號']),
      name: str(r['學生姓名']),
      tutor: str(r['導師姓名']),
      unit: str(r['單位'] || r['當月訓練科別'] || r['學校']),
      date: normDate(r['座談日期']),
    })).filter((r) => r.card && r.date);
  }

  // ---------- 名冊（統計檔） ----------
  function sheetCat(name) {
    if (/刪除/.test(name)) return null;
    if (/^R檢核/.test(name)) return 'R';
    if (/^PGY\(/.test(name)) return 'PGY';
    if (/^C\(/.test(name)) return 'UGY';
    if (/醫事人員/.test(name)) return '臨培';
    return null;
  }

  /**
   * sheets: [{ name, rows: 二維陣列 }] → { people: Map(card → 人), sheetsUsed }
   * PGY 有多個學年度工作表時，取起始年最大的那個。
   */
  function parseRoster(sheets) {
    const pick = {};
    for (const s of sheets) {
      const cat = sheetCat(s.name);
      if (!cat) continue;
      const start = Number((s.name.match(/\((\d+(?:\.\d+)?)/) || [])[1] || 0);
      if (!pick[cat] || start > pick[cat].start) pick[cat] = { ...s, start, cat };
    }
    const people = new Map();
    const sheetsUsed = {};
    for (const s of Object.values(pick)) {
      const hr = s.rows.slice(0, 3).findIndex((r) => (r || []).some((c) => /卡號/.test(str(c))));
      if (hr < 0) continue;
      const h = s.rows[hr].map(str);
      const col = (re) => h.findIndex((x) => re.test(x));
      const ci = {
        card: col(/卡號$/), name: col(/^(學員)?姓名$/), tutor: col(/^導師(姓名)?$/), unit: col(/^(單位|職類|學校)$/), gsm: col(/^(公務機|GSM)$/i),
      };
      sheetsUsed[s.cat] = s.name;
      s.rows.slice(hr + 1).forEach((r, i) => {
        const card = normCard((r || [])[ci.card]);
        if (!card || people.has(card)) return;
        people.set(card, {
          cat: s.cat, card, name: str(r[ci.name]), tutor: ci.tutor >= 0 ? str(r[ci.tutor]) : '',
          unit: ci.unit >= 0 ? str(r[ci.unit]) : '', gsm: ci.gsm >= 0 ? str(r[ci.gsm]).replace(/\D/g, '') : '', sheet: s.name, row: hr + 2 + i,
        });
      });
    }
    return { people, sheetsUsed };
  }

  /** 卡號 → 人（名冊優先，其次 rawdata 最新一筆） */
  function buildDirectory(roster, rawAll) {
    const dir = new Map(roster.people);
    [...rawAll].sort((a, b) => a.date.localeCompare(b.date)).forEach((r) => {
      const p = dir.get(r.card);
      if (!p) dir.set(r.card, { cat: r.cat, card: r.card, name: r.name, tutor: r.tutor, unit: r.unit, sheet: '', row: null });
      else if (!p.name) Object.assign(p, { name: r.name, tutor: p.tutor || r.tutor });
    });
    return dir;
  }

  // ---------- 簽到簿 ----------
  /** 由檔名判斷簽到簿類別 */
  function signinSheetCat(filename) {
    const f = str(filename);
    if (/青果森/.test(f)) return '青果森';
    if (/臨培/.test(f)) return '臨培';
    if (/實習|UGY|Intern/i.test(f)) return 'UGY';
    if (/PGY/i.test(f)) return 'PGY';
    if (/(^|[-_\s])R(\.|[-_\s]|$)|住院醫師/i.test(f)) return 'R';
    return '';
  }

  /** 'MM/DD'、'M/D'、'1020' → ISO（年取目標月份的年） */
  function parseSignDate(text, ym) {
    const s = str(text);
    const m = s.match(/^(\d{1,2})\s*[/\-.月]\s*(\d{1,2})日?$/) || s.match(/^(\d{2})(\d{2})$/);
    if (!m) return '';
    const y = Number(ym.slice(0, 4)); const mm = Number(m[1]); const dd = Number(m[2]);
    const d = new Date(y, mm - 1, dd);
    if (d.getMonth() !== mm - 1 || d.getDate() !== dd) return '';
    return `${y}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`;
  }

  // ---------- 計算 ----------
  /**
   * @param ym 'YYYY-MM'
   * @param signins [{ file, sheetCat, page, line, date(文字), card, name, illegible, note }]
   */
  function compute({ ym, rawAll, roster, signins, dupChoice = {} }) {
    const dir = buildDirectory(roster, rawAll);
    const catOf = (card) => dir.get(card)?.cat || guessCat(card);

    // 1) 本月 rawdata：每人取最早一筆
    const rawMonth = new Map();
    rawAll.filter((r) => r.date.startsWith(ym)).sort((a, b) => a.date.localeCompare(b.date)).forEach((r) => {
      const e = rawMonth.get(r.card);
      if (!e) rawMonth.set(r.card, { ...r, count: 1 });
      else e.count += 1;
    });

    // 2) 簽到：有效 / 無法辨識 / 重複
    const illegible = [];
    const valid = [];
    signins.forEach((s) => {
      const card = normCard(s.card);
      if (s.illegible || !card) { if (s.illegible || s.name || s.date) illegible.push(s); return; }
      valid.push({ ...s, card, name: dir.get(card)?.name || s.name, cat: catOf(card), isoDate: parseSignDate(s.date, ym) });
    });
    // 美德耐（非青果森）優先，其次青果森；同類依檔案、頁、列順序
    const ordered = [...valid].sort((a, b) => (a.sheetCat === '青果森') - (b.sheetCat === '青果森')
      || a.file.localeCompare(b.file) || a.page - b.page || a.line - b.line);
    // 重複申請：同卡號多筆 → 一組。dupChoice[card]：'all'（全部都算）或要保留那筆的 signKey；
    // 預設保留排序第一筆（美德耐優先）。
    const byCard = new Map();
    ordered.forEach((s) => { if (!byCard.has(s.card)) byCard.set(s.card, []); byCard.get(s.card).push(s); });
    const firstSign = new Map();
    const duplicates = [];
    const dupGroups = [];
    const dupSet = new Set();   // 被扣除的原始簽到物件（修正人數時排除）
    byCard.forEach((list, card) => {
      const choice = dupChoice[card];
      const kept = list.find((s) => signKey(s) === choice) || list[0];
      firstSign.set(card, kept);
      if (list.length < 2) return;
      dupGroups.push({ card, name: kept.name, cat: kept.cat, entries: list, keep: choice === 'all' ? 'all' : signKey(kept) });
      if (choice === 'all') return;
      list.filter((s) => s !== kept).forEach((s) => { dupSet.add(s); duplicates.push({ ...s, first: kept }); });
    });

    // 3) 餐費：原本人數（全部有效簽到）與修正人數（扣除重複）
    const group = (s) => {
      if (s.sheetCat !== '青果森') return s.sheetCat || catOf(s.card) || '未知';
      return catOf(s.card) === 'UGY' ? '青果森-UGY' : '青果森-R+PGY';
    };
    const MEAL_ROWS = [
      ['臨培', 180], ['UGY', 360], ['PGY', 180], ['R', 180], ['青果森-R+PGY', 180], ['青果森-UGY', 360],
    ];
    const meal = MEAL_ROWS.map(([g, rate]) => {
      const all = ordered.filter((s) => group(s) === g);
      const kept = all.filter((s) => !dupSet.has(s));
      return { group: g, rate, before: all.length, after: kept.length, amountBefore: all.length * rate, amountAfter: kept.length * rate };
    });
    const unknownGroup = ordered.filter((s) => !MEAL_ROWS.some(([g]) => g === group(s)));

    // 4) 檢核結果：名冊所有人 + 不在名冊但本月有紀錄/簽到的人
    const cards = new Set([...roster.people.keys(), ...rawMonth.keys(), ...firstSign.keys()]);
    const rows = [...cards].map((card) => {
      const p = dir.get(card) || { card, cat: guessCat(card), name: '', tutor: '', unit: '', sheet: '', row: null };
      const raw = rawMonth.get(card);
      const sign = firstSign.get(card);
      let status = '';
      if (sign) status = '院內';
      else if (raw) status = 'V';
      const red = !!sign && !raw && p.cat !== '臨培';
      const notes = [];
      if (!roster.people.has(card)) notes.push('名冊無此人');
      if (raw && raw.count > 1) notes.push(`本月系統 ${raw.count} 筆`);
      if (red) notes.push('有簽到但系統無紀錄');
      const signs = valid.filter((s) => s.card === card);
      if (signs.length > 1) notes.push(`簽到 ${signs.length} 次`);
      if (signs.some((x) => x.auto)) notes.push('手寫自動比對，請確認');
      return {
        ...p, name: p.name || sign?.name || '', status,
        date: raw?.date || sign?.isoDate || '',
        rawCount: raw?.count || 0,
        signSources: [...new Set(signs.map((s) => (s.sheetCat === '青果森' ? '青果森' : '美德耐')))].join('、'),
        red, notes,
      };
    });
    const catRank = (c) => { const i = CAT_ORDER.indexOf(c); return i < 0 ? 99 : i; };
    rows.sort((a, b) => catRank(a.cat) - catRank(b.cat) || (a.row ?? 1e9) - (b.row ?? 1e9) || a.card.localeCompare(b.card));

    return { ym, rows, meal, duplicates, dupGroups, illegible, unknownGroup, directory: dir, rawMonthCount: rawMonth.size };
  }

  /** 結報信件文字（仿「每月補助+重複申請」格式） */
  function mealMessage(result) {
    const rocYm = `${Number(result.ym.slice(0, 4)) - 1911}/${result.ym.slice(5, 7)}`;
    const m = Object.fromEntries(result.meal.map((x) => [x.group, x]));
    const fmt = (n) => n.toLocaleString('en-US');
    const line = (label, x) => (x.before === x.after
      ? `${label}：${x.after}人；計${fmt(x.amountAfter)}元無誤`
      : `${label}：${x.before}--->${x.after}人；計${fmt(x.amountBefore)}--->${fmt(x.amountAfter)}元`);
    const q = m['青果森-R+PGY']; const qu = m['青果森-UGY'];
    const qLine = `青果森：PGY+R ${q.before === q.after ? q.after : `${q.before}--->${q.after}`}人，UGY${qu.before === qu.after ? qu.after : `${qu.before}--->${qu.after}`}人；計${fmt(q.amountBefore + qu.amountBefore)}${q.before + qu.before === q.after + qu.after ? '' : `--->${fmt(q.amountAfter + qu.amountAfter)}`}元`;
    const lines = [rocYm, '您好：', '本月美德耐餐費補助結報如下：',
      line('臨培', m['臨培']), line('UGY', m.UGY), line('PGY', m.PGY), line('R', m.R), qLine,
      '再麻煩您開發票了，謝謝您。'];
    if (result.duplicates.length) {
      lines.push('', '以下已通知繳款：');
      result.duplicates.forEach((d) => lines.push(`${d.card}${d.name || ''}(${CAT_LABEL[d.cat] || d.cat || ''}) > ${d.sheetCat === '青果森' ? '青果森' : '美德耐'}`));
    }
    return lines.join('\n');
  }

  /** 產生結果活頁簿（ExcelJS 由呼叫端傳入） */
  function buildWorkbook(ExcelJS, result) {
    const wb = new ExcelJS.Workbook();
    const head = (ws, cols) => {
      ws.columns = cols.map(([header, width]) => ({ header, width }));
      ws.getRow(1).font = { bold: true };
      ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F3F7' } };
      ws.views = [{ state: 'frozen', ySplit: 1 }];
    };
    const toDate = (iso) => (iso ? new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10))) : null);

    const ws = wb.addWorksheet('檢核結果');
    head(ws, [['類別', 8], ['名冊工作表', 22], ['名冊列號', 9], ['單位', 14], ['卡號', 9], ['姓名', 10], ['導師', 10],
      ['狀態', 8], ['座談日期', 12], ['系統筆數', 9], ['簽到來源', 14], ['備註', 30]]);
    result.rows.forEach((r) => {
      const row = ws.addRow([CAT_LABEL[r.cat] || r.cat, r.sheet, r.row, r.unit, r.card, r.name, r.tutor,
        r.status, toDate(r.date), r.rawCount || null, r.signSources, r.notes.join('；')]);
      row.getCell(9).numFmt = 'yyyy/mm/dd';
      if (r.red) {
        const c = row.getCell(8);
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
        c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      }
    });
    ws.autoFilter = { from: 'A1', to: 'L1' };

    const wm = wb.addWorksheet('餐費補助');
    head(wm, [['職類', 16], ['原本人數', 10], ['修正人數', 10], ['單價', 8], ['原本金額', 12], ['修正金額', 12]]);
    result.meal.forEach((x) => wm.addRow([x.group, x.before, x.after, x.rate, x.amountBefore, x.amountAfter]));
    const n = result.meal.length + 1;
    wm.addRow(['總和', { formula: `SUM(B2:B${n})` }, { formula: `SUM(C2:C${n})` }, null,
      { formula: `SUM(E2:E${n})` }, { formula: `SUM(F2:F${n})` }]).font = { bold: true };
    ['E', 'F'].forEach((c) => { wm.getColumn(c).numFmt = '#,##0'; });
    wm.addRow([]);
    mealMessage(result).split('\n').forEach((l) => wm.addRow([l]));

    const wd = wb.addWorksheet('重複申請');
    head(wd, [['卡號', 9], ['姓名', 10], ['類別', 12], ['簽到簿', 24], ['頁', 5], ['列', 5], ['日期', 8], ['處理', 12]]);
    result.dupGroups.forEach((g) => {
      g.entries.forEach((s) => {
        let action = '全部計入';
        if (g.keep !== 'all') action = signKey(s) === g.keep ? '保留' : '重複（扣除）';
        const row = wd.addRow([g.card, g.name, CAT_LABEL[g.cat] || g.cat, s.file, s.page, s.line, s.date, action]);
        if (action === '重複（扣除）') row.getCell(8).font = { color: { argb: 'FFFF0000' }, bold: true };
      });
      wd.addRow([]);
    });

    const wi = wb.addWorksheet('無法辨識');
    head(wi, [['檔案', 24], ['頁', 5], ['列', 5], ['日期', 8], ['卡號', 9], ['姓名', 10], ['說明', 30]]);
    result.illegible.forEach((s) => { wi.addRow([s.file, s.page, s.line, s.date, s.card, s.name, s.note || '無法辨識']).font = { color: { argb: 'FFFF0000' } }; });
    return wb;
  }

  return {
    RATES, CAT_ORDER, CAT_LABEL, signKey, normCard, normDate, guessCat, detectRawType, parseRaw, parseRoster,
    buildDirectory, signinSheetCat, parseSignDate, compute, mealMessage, buildWorkbook,
  };
});
