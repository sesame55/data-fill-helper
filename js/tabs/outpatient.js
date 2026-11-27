/*
 * 分頁：教學門住診 → 教學門診
 * 流程：匯入（範本 / rawdata / 各單位申請單 .doc）→ 申請單比對 → 計算 → 下載填好的名冊
 * 所有資料只存在記憶體（state），清空、關閉分頁即消失。
 */
(() => {
  const O = window.OutpatientCalc;
  const Fill = window.InpatientFill;
  const DocReader = window.DocReader;

  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = (n) => n.toLocaleString('en-US');
  const mmdd = (iso) => `${iso.slice(5, 7)}/${iso.slice(8, 10)}`;

  const emptyState = () => ({
    template: null,   // { name, buffer, info }
    raw: null,        // { name, rows }
    parsed: null,     // O.parseRaw 結果
    rawRows: [],      // AppForm.parseRawRows（PDF 申請單比對用）
    unclaimed: [],
    apps: [],         // [{ file, ...O.parseApplication 結果 }]
    cardOverride: {}, // 申請單老師姓名 → 手動卡號
    appRecords: [],
    compare: [],
    summary: [],
    busy: [],
  });
  let state = emptyState();

  // ---------- 匯入 ----------
  const dropzone = $('op-dropzone');
  const input = $('op-file');
  input.addEventListener('change', () => { handleFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
  dropzone.addEventListener('drop', (e) => handleFiles([...e.dataTransfer.files]));

  async function handleFiles(files) {
    for (const f of files) {
      const ext = f.name.split('.').pop().toLowerCase();
      setBusy(f.name, true);
      try {
        if (ext === 'doc') await loadDoc(f);
        else if (ext === 'pdf') { $('op-app-step').hidden = false; await appPdf.addPdf(f); }
        else if (ext === 'xlsx') await loadXlsx(f);
        else if (ext === 'docx') alert(`「${f.name}」是 .docx，目前請在 Word 另存為「Word 97-2003 文件 (.doc)」後再匯入。`);
        else alert(`不支援的檔案：${f.name}`);
      } catch (err) {
        console.error(err);
        alert(`讀取「${f.name}」失敗：${err.message}`);
      } finally {
        setBusy(f.name, false);
      }
    }
    guessMonth(files);
    recompute();
  }

  /** 沒有範本時，從檔名猜月份：「202609」或民國「11509」；多個檔案時取出現最多的月份 */
  function guessMonth(files) {
    const el = $('op-month');
    if (state.template) { el.value = targetYm(); return; }
    if (el.value) return;
    const cnt = {};
    files.forEach((f) => {
      const w = f.name.match(/20(\d{2})(0[1-9]|1[0-2])/);
      const r = f.name.match(/1(\d{2})(0[1-9]|1[0-2])(?!\d)/);
      const ym = w ? `20${w[1]}-${w[2]}` : r ? `${Number(`1${r[1]}`) + 1911}-${r[2]}` : '';
      if (ym) cnt[ym] = (cnt[ym] || 0) + 1;
    });
    const best = Object.entries(cnt).sort((x, y) => y[1] - x[1] || y[0].localeCompare(x[0]))[0];
    if (best) el.value = best[0];
  }

  async function loadXlsx(file) {
    const buffer = await file.arrayBuffer();
    const names = XLSX.read(buffer, { bookSheets: true }).SheetNames;
    if (names.includes('課程rawdata')) {
      const wb = XLSX.read(buffer, { sheets: ['課程rawdata'], dense: true });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets['課程rawdata'], { defval: '', raw: false });
      if (rows.length && !String(rows[0]['課程'] || '').includes('門診')) {
        if (!confirm(`「${file.name}」看起來不是教學門診的 rawdata（課程：${rows[0]['課程']}），仍要使用嗎？`)) return;
      }
      state.raw = { name: file.name, rows, buffer };
      state.rawRows = AppForm.parseRawRows(rows);
      return;
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    const info = Fill.analyzeTemplate(ws);
    // 已填好卡號的名冊（例如上個月的正確版）→ 當作「單位對照」：學習 卡號 → 單位
    const ref = new Map();
    for (let r = info.startRow; r < info.totalRow; r += 1) {
      const row = ws.getRow(r);
      const card = String(row.getCell(info.cols.card).text || '').trim();
      const unit = info.cols.unit ? String(row.getCell(info.cols.unit).text || '').trim() : '';
      if (card && unit) ref.set(card.padStart(4, '0'), unit);
    }
    if (ref.size) { state.unitRef = ref; state.unitRefName = file.name; return; }
    if (!info.rocYear || !info.month) throw new Error('範本第一筆資料列找不到「年」「月」');
    state.template = { name: file.name, buffer, info };
  }

  async function loadDoc(file) {
    const doc = DocReader.parse(new Uint8Array(await file.arrayBuffer()));
    const year = targetYm() ? Number(targetYm().slice(0, 4)) : new Date().getFullYear();
    const app = O.parseApplication(doc, year);
    state.apps = state.apps.filter((a) => a.file !== file.name);
    state.apps.push({ file: file.name, ...app });
  }

  function setBusy(name, on) {
    state.busy = on ? [...state.busy, name] : state.busy.filter((n) => n !== name);
    renderStatus();
  }

  // ---------- 計算 ----------
  // ---------- 申請單（PDF）：計費基準 ----------
  const targetYm = () => (state.template ? `${state.template.info.rocYear + 1911}-${String(state.template.info.month).padStart(2, '0')}` : ($('op-month').value || ''));
  /** 名冊資訊：有範本用範本；沒有範本用月份欄 */
  const tinfo = () => {
    if (state.template) return state.template.info;
    const ym = targetYm();
    return { rocYear: Number(ym.slice(0, 4)) - 1911, month: Number(ym.slice(5, 7)), capacity: Infinity };
  };
  const appPdf = AppPdf.create({
    root: $('op-app'),
    getRawRows: () => state.rawRows || [],
    getYm: targetYm,
    onChange: () => recompute(true),
  });

  function recompute(fromApp) {
    const ym0 = targetYm();
    state.parsed = state.raw && ym0 ? O.parseRaw(state.raw.rows, Number(ym0.slice(0, 4)), Number(ym0.slice(5, 7))) : null;
    // 沒有申請單時的排序依據：rawdata 原始順序
    if (state.parsed) state.parsed.records.forEach((r, i) => { r.order = 5e8 + i; });
    if (!fromApp && appPdf.hasData()) { appPdf.rematch(); return; }   // rematch 完會再呼叫 recompute(true)
    if (state.parsed) {
      const prefix = targetYm();
      state.appRecords = O.applicationRecords(state.apps, state.parsed.directory, prefix, state.cardOverride);
      state.compare = O.compare(state.parsed.records, state.appRecords);
      const hasApp = state.apps.length > 0 || appPdf.hasData();
      // 計費規則（依 115/09 正確版歸納，2026-11-17）：
      //  1. 一般內科（Word 申請單上的老師）：該老師當月系統「全部」已評核紀錄。
      //  2. 其他單位（PDF 申請單）：只算申請單上「核准」的列——Word 版全部；系統印出版頁上有審核畫記「/」
      //     只算有畫記的列，沒有畫記則看申請單上「老師完成否」欄是否為 Y（不看系統現況）。
      //  3. 每位學員 500 元，每位老師每日上限 3 位。排序依申請單出現順序，一般內科在最後。
      const special = new Set(state.appRecords.map((r) => r.card).filter(Boolean));
      // 一般內科：依 Word 申請單上老師出現的順序排在最後
      const docOrder = new Map();
      state.appRecords.forEach((r, i) => { if (r.card && !docOrder.has(r.card)) docOrder.set(r.card, i); });
      const specialRecs = state.parsed.records.filter((r) => special.has(r.card))
        .map((r) => ({ ...r, inSystem: true, appUnit: '一般內科', order: 1e9 + (docOrder.get(r.card) ?? 0) }));
      // 內科部門診表（塗色格）：之後頁面的內科部申請單，只算塗色格上的老師；單位依科別代碼
      const sched = appPdf.schedule();
      const schedCard = new Map();   // 卡號 → 單位
      // 塗色格下方讀到的姓名（可能多讀了旁邊的字）→ 找 rawdata 中最相似的老師（含部分比對）
      const teachers = new Map();
      state.rawRows.forEach((r) => { if (r.tName && r.tName.length >= 2 && !teachers.has(r.tCard)) teachers.set(r.tCard, r.tName); });   // 略過姓名空白的紀錄
      const findTeacher = (txt) => {
        let best = null; let bs = 0;
        teachers.forEach((name, card) => {
          const sc = Math.max(AppForm.sim(txt, name), txt.includes(name) ? 1 : 0, ...[0, 1, 2].map((i) => AppForm.sim(txt.slice(i, i + name.length), name)));
          if (sc > bs) { bs = sc; best = card; }
        });
        return bs >= 0.6 ? best : '';
      };
      state.schedule = sched.map((x) => ({ ...x, card: findTeacher(x.teacher) }));
      state.schedule.forEach((x) => { if (x.card && O.CODE_UNIT[x.code] && x.code !== 'GM') schedCard.set(x.card, O.CODE_UNIT[x.code]); });
      // 每個檔案取塗色格最多的那一頁當排班表；之後的頁面為內科部申請單
      const schedPage = new Map();
      [...new Set(sched.map((x) => x.file))].forEach((f) => {
        const cnt = {}; sched.filter((x) => x.file === f).forEach((x) => { cnt[x.page] = (cnt[x.page] || 0) + 1; });
        schedPage.set(f, Number(Object.entries(cnt).sort((a, b) => b[1] - a[1])[0][0]));
      });
      const isInternal = (r) => schedPage.has(r.file) && r.page > schedPage.get(r.file);
      const ok = (r) => r.eligible !== false && !special.has(r.card) && (!isInternal(r) || schedCard.has(r.card));
      const monthRecs = appPdf.records().filter((r) => r.date.startsWith(prefix));
      const pdfRecs = monthRecs.filter(ok).map((r) => ({ ...r, student: O.normName(r.student) }));
      state.skippedRows = monthRecs.filter((r) => !ok(r) && !special.has(r.card));
      // 沒有申請單 → 暫以系統全部已評核紀錄
      let base = hasApp ? [...pdfRecs, ...specialRecs] : state.parsed.records.map((r) => ({ ...r, inSystem: true }));
      if (!$('op-include-app').checked) base = base.filter((r) => r.inSystem !== false);
      const all = O.summarize(base);
      const sysOnly = O.summarize(base.filter((r) => r.inSystem !== false));
      state.summary = O.markApplicationEffect(all, sysOnly);
      // 單位：過去的正確版名冊（卡號→單位）優先；一般內科固定；其餘用申請單／系統
      const key = (x) => x.card || `name:${x.name}`;
      state.summary.forEach((x) => {
        if (special.has(x.card)) x.unit = '一般內科';
        else if (schedCard.has(x.card)) x.unit = schedCard.get(x.card);
        else if (state.unitRef?.has(x.card)) x.unit = state.unitRef.get(x.card);
        else x.unit = O.officialUnit(x.unit);
      });
      // 排序：申請單出現順序；一般內科最後
      const firstOrder = new Map();
      base.forEach((r) => { const k = key(r); const o = r.order ?? 1e8; if (!firstOrder.has(k) || o < firstOrder.get(k)) firstOrder.set(k, o); });
      // 內科部：依門診表塗色格的時間順序，排在其他單位之後、一般內科之前
      state.schedule.forEach((x) => {
        if (x.card && schedCard.has(x.card) && !special.has(x.card)) {
          const o = 1e7 + (x.seq ?? 0);
          if (!firstOrder.has(x.card) || firstOrder.get(x.card) < 1e7 || o < firstOrder.get(x.card)) firstOrder.set(x.card, o);
        }
      });
      state.summary.sort((a, b) => (firstOrder.get(key(a)) ?? 1e8) - (firstOrder.get(key(b)) ?? 1e8));
      // 系統有紀錄、但不在任何申請單上（參考用）
      state.unclaimed = hasApp
        ? state.parsed.records.filter((x) => !base.some((r) => r.card === x.card && r.date === x.date && O.sameName(x.student, r.student)))
        : [];
      // 日期無法辨識的申請單列：紅字、金額空白（下載後人工核對）
      if (appPdf.hasData()) {
        appPdf.undated().forEach((u) => state.summary.push({
          card: '', name: u.teacher || `（無法辨識：${u.where}）`, unit: u.unit, days: [], amount: null,
          highlight: true, appExtra: 0, fromApp: 0, overCap: 0, daysText: `日期無法辨識（${u.where}）`,
        }));
      }
    } else {
      state.appRecords = []; state.compare = []; state.summary = []; state.unclaimed = [];
    }
    renderStatus();
    renderCompare();
    renderResult();
  }


  // ---------- 畫面 ----------
  function renderStatus() {
    const t = state.template; const r = state.raw; const p = state.parsed;
    const item = (ok, label, detail) =>
      `<div class="status-item ${ok ? 'ok' : ''}"><span class="dot"></span><b>${label}</b><span class="muted small">${detail}</span></div>`;
    const appDetail = state.apps.length
      ? state.apps.map((a) => {
        const mismatch = t && a.month && (a.month !== t.info.month || a.rocYear !== t.info.rocYear);
        return `${esc(a.unit || a.file)}（${a.entries.length} 列${a.errors.length ? `，<span class="problem">${a.errors.length} 個問題</span>` : ''}${mismatch ? `，<span class="problem">月份 ${a.rocYear}/${a.month} 與範本不同</span>` : ''}）`;
      }).join('、')
      : '（選填，可多個）';
    $('op-status').innerHTML = [
      item(!!t, '範本', t ? `${esc(t.name)}・${t.info.rocYear} 年 ${t.info.month} 月・可填 ${t.info.capacity} 列` : '（選填）未匯入時，依「月份」產生同格式的空白名冊'),
      item(!!r, 'rawdata', r ? `${esc(r.name)}・${money(r.rows.length)} 筆${p ? `・本月已評核 ${money(p.records.length)} 筆` : ''}` : '尚未匯入'),
      item(state.apps.length > 0, '申請單', appDetail),
      item(!!state.unitRef, '單位對照', state.unitRef ? `${esc(state.unitRefName)}・${state.unitRef.size} 位老師` : '（選填：匯入過去已填好的名冊，單位會依它填寫）'),
      state.busy.length ? `<div class="status-item busy"><span class="dot"></span>讀取中：${esc(state.busy.join('、'))}</div>` : '',
    ].join('');
  }

  const STATUS_CLASS = { 一致: 'ok', 有差異: 'diff', 系統無紀錄: 'bad', 申請單未列: 'diff' };

  function renderCompare() {
    const show = state.parsed && state.apps.length > 0;
    $('op-cmp-step').hidden = !show;
    if (!show) return;
    const rows = state.compare;
    const count = (s) => rows.filter((r) => r.status === s).length;
    $('op-cmp-summary').innerHTML = `
      <div class="stat"><span>一致</span><b>${count('一致')}</b></div>
      <div class="stat"><span>有差異</span><b>${count('有差異')}</b></div>
      <div class="stat"><span>系統無紀錄</span><b>${count('系統無紀錄')}</b></div>
      <div class="stat"><span>申請單未列</span><b>${count('申請單未列')}</b></div>`;
    const diffOnly = $('op-cmp-diff-only').checked;
    const list = diffOnly ? rows.filter((r) => r.status !== '一致') : rows;
    $('op-cmp').innerHTML = list.length ? list.map((r) => `
      <tr>
        <td>${esc(r.unit)}</td><td>${esc(r.name)}</td>
        <td>${r.card ? esc(r.card) : `<input class="card-input" data-teacher="${esc(r.name)}" placeholder="請填卡號" size="6">`}</td>
        <td>${mmdd(r.date)}</td><td>${r.appCount}</td><td>${r.sysCount}</td>
        <td class="text-red">${esc(r.onlyApp.join('、'))}</td><td>${esc(r.onlySys.join('、'))}</td>
        <td><span class="badge ${STATUS_CLASS[r.status]}">${r.status}</span></td>
      </tr>`).join('') : '<tr class="empty"><td colspan="9">全部一致</td></tr>';
  }

  $('op-cmp').addEventListener('change', (e) => {
    const inp = e.target.closest('.card-input');
    if (!inp) return;
    state.cardOverride[inp.dataset.teacher] = inp.value.trim();
    recompute();
  });
  $('op-cmp-diff-only').addEventListener('change', renderCompare);
  $('op-include-app').addEventListener('change', recompute);
  $('op-month').addEventListener('change', () => recompute());

  function renderResult() {
    const show = !!state.parsed;
    $('op-result-step').hidden = !show;
    if (!show) return;
    const s = state.summary;
    const total = s.reduce((a, x) => a + (x.amount || 0), 0);
    const sysOnlyTotal = O.summarize(state.parsed.records).reduce((a, x) => a + x.amount, 0);
    const redCount = s.filter((x) => x.highlight).length;
    $('op-stats').innerHTML = `
      <div class="stat"><span>老師人數</span><b>${s.length}</b></div>
      <div class="stat"><span>總金額</span><b>${money(total)}</b></div>
      <div class="stat muted-stat"><span>僅系統資料時</span><b>${money(sysOnlyTotal)}</b></div>
      <div class="stat"><span>紅字（申請單增加金額）</span><b class="text-red">${redCount} 位</b></div>
      <div class="stat"><span>系統紀錄</span><b>${money(state.parsed.records.length)}</b></div>
      <div class="stat muted-stat"><span>未評核完成（不計）</span><b>${money(state.parsed.skipped.notDone)}</b></div>`;

    const warns = [];
    if (!state.apps.length && !appPdf.hasData()) warns.push('尚未匯入申請單：目前暫以 rawdata 全部已評核紀錄計算（正式計費請匯入申請單）。');
    if (appPdf.busy()) warns.push('申請單辨識中…結果會自動更新。');
    if (state.skippedRows?.length) {
      const who = [...new Set(state.skippedRows.map((r) => r.name))];
      warns.push(`申請單上不計的列（申請單「老師完成否」空白，或內科部門診表上沒有塗色）：${state.skippedRows.length} 列（${who.slice(0, 12).map(esc).join('、')}${who.length > 12 ? '…' : ''}）`);
    }
    const noCardApp = s.filter((x) => !x.card && x.highlight).length;
    if (noCardApp) warns.push(`有 ${noCardApp} 列為申請單上無法辨識或找不到卡號的老師，下載的名冊中為<b>紅字、卡號空白</b>（日期無法辨識者金額也空白），請人工核對。`);
    if (state.unclaimed.length) {
      const tea = [...new Set(state.unclaimed.map((r) => r.name))];
      warns.push(`系統有紀錄、但申請單沒有列的：${state.unclaimed.length} 筆（${tea.length} 位老師，不計費）：${tea.slice(0, 15).map(esc).join('、')}${tea.length > 15 ? '…' : ''}`);
    }
    const noCard = s.filter((x) => !x.card);
    state.apps.filter((a) => a.errors.length).forEach((a) => warns.push(`${esc(a.unit || a.file)} 申請單：${a.errors.map(esc).join('；')}`));
    if (!state.template) warns.push('未匯入範本：下載時會產生同格式的空白名冊（年、月、單位、卡號、姓名、金額、備註）。');
    else if (s.length > tinfo().capacity) {
      warns.push(`人數 ${s.length} 超過範本 ${tinfo().capacity} 列，下載時會自動增加列數並更新合計公式`);
    }
    $('op-warnings').innerHTML = warns.map((w) => `<div class="warn">⚠ ${w}</div>`).join('');

    const q = $('op-filter').value.trim();
    const rows = q ? s.filter((x) => `${x.unit} ${x.card} ${x.name}`.includes(q)) : s;
    // 與下載的名冊相同：年｜月｜單位｜卡號｜姓名｜金額｜備註；紅字列同 Excel；滑鼠停留可看教學日與說明
    const ti = tinfo();
    $('op-result').innerHTML = rows.length ? rows.map((x) => {
      const notes = [x.days.length ? O.formatDays(x.days) : (x.daysText || '')];
      if (x.highlight) notes.push(x.amount === null ? '日期無法辨識' : !x.card ? '無法辨識／缺卡號' : `系統無紀錄 +${money(x.appExtra)}`);
      if (x.overCap) notes.push(`${x.overCap} 天超過上限`);
      return `<tr class="${x.highlight ? 'row-red' : ''}" title="${esc(notes.filter(Boolean).join('；'))}">
        <td>${ti.rocYear}</td><td>${ti.month}</td><td>${esc(x.unit)}</td><td>${esc(x.card)}</td><td>${esc(x.name)}</td>
        <td class="num">${x.amount === null ? '' : money(x.amount)}</td><td>匯入銀行帳號</td></tr>`;
    }).join('') + `<tr class="total-row"><td colspan="4"></td><td>合計</td><td class="num">${money(total)}</td><td></td></tr>`
      : '<tr class="empty"><td colspan="7">沒有符合的資料</td></tr>';
  }
  $('op-filter').addEventListener('input', renderResult);

  // ---------- 下載比對後的 rawdata ----------
  $('op-raw-download').addEventListener('click', async () => {
    if (!state.raw || (!appPdf.hasData() && !state.apps.length)) { alert('請先匯入 rawdata 與申請單'); return; }
    // PDF 申請單對到的列 ＋ Word 申請單在系統有紀錄的列
    const rows = [...appPdf.matchedRowNums()];
    (state.appRecords || []).forEach((r) => {
      const hit = state.rawRows.find((x) => x.tCard === r.card && x.date === r.date && O.sameName(x.sName, r.student));
      if (hit?.rowNum) rows.push(hit.rowNum);
    });
    try {
      await AppPdf.downloadMatchedRaw(state.raw, rows);
    } catch (err) { console.error(err); alert(`產生檔案失敗：${err.message}`); }
  });

  // ---------- 下載 ----------
  $('op-download').addEventListener('click', async () => {
    if (!state.summary.length) return;
    const btn = $('op-download');
    btn.disabled = true;
    try {
      // 有範本填範本；沒有範本產生同格式的空白名冊
      let wb;
      if (state.template) { wb = new ExcelJS.Workbook(); await wb.xlsx.load(state.template.buffer.slice(0)); } else {
        const ti = tinfo();
        wb = Fill.blankTemplate(ExcelJS, { title: '教學門診指導費支領名冊', rocYear: ti.rocYear, month: ti.month, withDays: false });
      }
      Fill.fillTemplate(wb, state.summary);
      const out = await wb.xlsx.writeBuffer();
      const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = state.template ? `${state.template.name.replace(/\.xlsx$/i, '')}_已填寫.xlsx` : `教學門診名冊_${tinfo().rocYear}${String(tinfo().month).padStart(2, '0')}.xlsx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error(err);
      alert(`產生檔案失敗：${err.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // ---------- 清空 ----------
  App.registerTab('teaching-clinic-outpatient', {
    reset() {
      state = emptyState();
      appPdf.reset();
      $('op-app-step').hidden = true;
      $('op-filter').value = '';
      recompute();
    },
  });
  renderStatus();
})();
