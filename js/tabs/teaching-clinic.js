/*
 * 分頁：教學門住診 → 教學住診
 * 流程：匯入（範本 / rawdata / 手寫 PDF）→ 確認手寫紀錄 → 計算 → 下載填好的名冊
 * 所有資料只存在記憶體（state），登出、清空、關閉分頁即消失。
 */
(() => {
  const Calc = window.InpatientCalc;
  const Fill = window.InpatientFill;
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = (n) => n.toLocaleString('en-US');

  const emptyState = () => ({
    template: null,   // { name, buffer, info }
    raw: null,        // { name, rows }
    parsed: null,     // Calc.parseRaw 結果
    pages: [],        // 手寫紀錄：{ id, file, pageNo, cropUrl, fullUrl, teacher, card, unit, date, include }
    summary: [],
    unclaimed: [],
    rawRows: [],
    busy: [],
  });
  let state = emptyState();
  let pageSeq = 0;

  // ---------- 匯入 ----------
  const dropzone = $('ip-dropzone');
  const input = $('ip-file');

  input.addEventListener('change', () => { handleFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
  dropzone.addEventListener('drop', (e) => handleFiles([...e.dataTransfer.files]));

  async function handleFiles(files) {
    for (const f of files) {
      const ext = f.name.split('.').pop().toLowerCase();
      setBusy(f.name, true);
      try {
        // 所有 PDF（補助費申請單、手寫住診教學紀錄表）一律自動辨識；辨識不出的在名冊標紅字
        if (ext === 'pdf') { $('ip-app-step').hidden = false; await app.addPdf(f); }
        else if (ext === 'xlsx') await loadXlsx(f);
        else alert(`不支援的檔案：${f.name}`);
      } catch (err) {
        console.error(err);
        alert(`讀取「${f.name}」失敗：${err.message}`);
      } finally {
        setBusy(f.name, false);
      }
    }
    guessMonth(files);
    reparse();
  }

  /** 沒有範本時，從檔名猜月份：「202609」或民國「11509」；多個檔案時取出現最多的月份 */
  function guessMonth(files) {
    const el = $('ip-month');
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
      state.raw = { name: file.name, rows, buffer };
      state.rawRows = AppForm.parseRawRows(rows);
      return;
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const info = Fill.analyzeTemplate(wb.worksheets[0]);
    if (!info.rocYear || !info.month) throw new Error('範本第一筆資料列找不到「年」「月」');
    state.template = { name: file.name, buffer, info };
  }

  async function loadPdf(file) {
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false }).promise;
    for (let p = 1; p <= pdf.numPages; p += 1) {
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

      // 表頭區（科別、指導教師、教學日期、學員）約在頁面上方 2%～22%
      const crop = document.createElement('canvas');
      const sy = Math.round(canvas.height * 0.02);
      const sh = Math.round(canvas.height * 0.20);
      crop.width = canvas.width;
      crop.height = sh;
      crop.getContext('2d').drawImage(canvas, 0, sy, canvas.width, sh, 0, 0, canvas.width, sh);

      state.pages.push({
        id: ++pageSeq, file: file.name, pageNo: p,
        cropUrl: crop.toDataURL('image/jpeg', 0.85),
        fullUrl: canvas.toDataURL('image/jpeg', 0.85),
        teacher: '', card: '', unit: '', date: '', include: true,
      });
      canvas.width = 0; crop.width = 0;   // 釋放記憶體
    }
    pdf.destroy();
  }

  function setBusy(name, on) {
    state.busy = on ? [...state.busy, name] : state.busy.filter((n) => n !== name);
    renderStatus();
  }

  // ---------- 解析 / 計算 ----------
  function reparse() {
    const ym = targetYm();
    state.parsed = state.raw && ym ? Calc.parseRaw(state.raw.rows, Number(ym.slice(0, 4)), Number(ym.slice(5, 7))) : null;
    // 沒有申請單時的排序依據：rawdata 原始順序
    if (state.parsed) state.parsed.records.forEach((r, i) => { r.order = 5e8 + i; });
    renderStatus();
    renderPages();
    if (app.hasData()) app.rematch(); else recalc();
  }

  // ---------- 申請單（PDF）：計費基準 ----------
  const targetYm = () => (state.template ? `${state.template.info.rocYear + 1911}-${String(state.template.info.month).padStart(2, '0')}` : ($('ip-month').value || ''));
  /** 名冊資訊：有範本用範本；沒有範本用月份欄 */
  const tinfo = () => {
    if (state.template) return state.template.info;
    const ym = targetYm();
    return { rocYear: Number(ym.slice(0, 4)) - 1911, month: Number(ym.slice(5, 7)), capacity: Infinity };
  };
  const app = AppPdf.create({
    root: $('ip-app'),
    getRawRows: () => state.rawRows || [],
    getYm: targetYm,
    onChange: () => recalc(),
  });

  /** 'MM/DD'、'M/D'、'0905' → 'YYYY-MM-DD'（年取範本年份）；格式錯誤回傳 '' */
  function parsePageDate(text) {
    const t = state.template?.info;
    if (!t) return '';
    const s = String(text).trim();
    const m = s.match(/^(\d{1,2})\s*[/\-.月]\s*(\d{1,2})日?$/) || s.match(/^(\d{2})(\d{2})$/);
    if (!m) return '';
    const mm = Number(m[1]); const dd = Number(m[2]);
    const y = t.rocYear + 1911;
    const d = new Date(y, mm - 1, dd);
    if (d.getMonth() !== mm - 1 || d.getDate() !== dd) return '';
    return `${y}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`;
  }

  function pageProblem(p) {
    const t = state.template?.info;
    if (!p.include) return '';
    if (!p.teacher.trim()) return '請填指導教師';
    const iso = parsePageDate(p.date);
    if (!iso) return '教學日格式請用 09/05';
    if (t && Number(iso.slice(5, 7)) !== t.month) return `教學日不是 ${t.month} 月`;
    if (!p.card.trim()) return '缺卡號（可先計算，下載前請補上）';
    return '';
  }

  function recalc() {
    if (!state.parsed) { $('ip-result-step').hidden = true; return; }
    const manual = state.pages
      .filter((p) => p.include && p.teacher.trim() && parsePageDate(p.date)
        && Number(parsePageDate(p.date).slice(5, 7)) === tinfo().month)
      .map((p) => ({
        card: Calc.normCard(p.card), name: p.teacher.trim(), unit: p.unit.trim(),
        date: parsePageDate(p.date), source: '手寫',
      }));
    const weekStart = Number($('ip-week-start').value);
    // 計費基準：有申請單 → 只算申請單上的列（系統沒有的紅字計入）；沒有申請單 → 暫以系統全部已評核紀錄
    const ym = targetYm();
    const base = app.hasData()
      ? app.records().filter((r) => r.date.startsWith(ym)).map((r) => ({ ...r, unit: r.appUnit || r.unit }))
      : state.parsed.records;
    state.summary = Calc.summarize([...base, ...manual], { weekStart });
    const sysOnly = Calc.summarize([...base.filter((r) => r.inSystem !== false), ...manual], { weekStart });
    const sysAmt = new Map(sysOnly.map((x) => [x.card || `name:${x.name}`, x.amount]));
    state.summary.forEach((x) => {
      x.appExtra = x.amount - (sysAmt.get(x.card || `name:${x.name}`) || 0);
      x.highlight = x.appExtra > 0;
    });
    // 系統有紀錄、但沒有出現在申請單（參考用，不計費）
    if (app.hasData()) {
      const claimed = new Set(app.matched().filter((m) => m.match).map((m) => `${m.match.tCard}|${m.match.date}`));
      state.unclaimed = state.parsed.records.filter((r) => !claimed.has(`${r.card}|${r.date}`));
    } else state.unclaimed = [];
    // 排序：依彙整表（申請單）出現順序，不依筆畫；一般內科放最後
    const key = (x) => x.card || `name:${x.name}`;
    const first = new Map();
    base.forEach((r) => { const k = key(r); const o = r.order ?? 9e8; if (!first.has(k) || o < first.get(k)) first.set(k, o); });
    const rank = (x) => (Calc.LAST_UNITS.includes(x.unit) ? 1e12 : 0) + (first.get(key(x)) ?? 9e8);
    state.summary.sort((a, b) => rank(a) - rank(b));
    // 日期無法辨識的申請單列：紅字、金額空白（下載後人工核對）
    if (app.hasData()) {
      app.undated().forEach((u) => state.summary.push({
        card: '', name: u.teacher || `（無法辨識：${u.where}）`, unit: u.unit, days: [], weeks: [],
        amount: null, highlight: true, appExtra: 0, daysText: `日期無法辨識（${u.where}）`,
      }));
    }
    renderResult();
  }

  // ---------- 畫面 ----------
  function renderStatus() {
    const t = state.template; const r = state.raw; const p = state.parsed;
    const pdfFiles = [...new Set(state.pages.map((x) => x.file))];
    const item = (ok, label, detail) =>
      `<div class="status-item ${ok ? 'ok' : ''}"><span class="dot"></span><b>${label}</b><span class="muted small">${detail}</span></div>`;
    $('ip-status').innerHTML = [
      item(!!t, '範本', t ? `${esc(t.name)}・${t.info.rocYear} 年 ${t.info.month} 月・可填 ${t.info.capacity} 列` : '（選填）未匯入時，依「月份」產生同格式的空白名冊'),
      item(!!r, 'rawdata', r ? `${esc(r.name)}・${money(r.rows.length)} 筆${p ? `・本月已評核 ${money(p.records.length)} 筆` : ''}` : '尚未匯入'),
      item(pdfFiles.length > 0, '手寫紀錄', pdfFiles.length ? `${pdfFiles.length} 個檔案・${state.pages.length} 頁` : '（選填）'),
      state.busy.length ? `<div class="status-item busy"><span class="dot"></span>讀取中：${esc(state.busy.join('、'))}</div>` : '',
    ].join('');
  }

  function renderPages() {
    $('ip-pdf-step').hidden = state.pages.length === 0;
    const names = state.parsed ? [...state.parsed.directory.keys()].sort() : [];
    $('ip-pdf-list').innerHTML = `
      <datalist id="ip-teacher-names">${names.map((n) => `<option value="${esc(n)}">`).join('')}</datalist>
      ${state.pages.map((p) => `
        <div class="pdf-card ${p.include ? '' : 'excluded'}" data-id="${p.id}">
          <img src="${p.cropUrl}" alt="第 ${p.pageNo} 頁表頭" data-full="${p.id}" title="點擊看整頁">
          <div class="pdf-fields">
            <div class="muted small">${esc(p.file)}・第 ${p.pageNo} 頁</div>
            <label>指導教師<input data-f="teacher" list="ip-teacher-names" value="${esc(p.teacher)}"></label>
            <label>教學日<input data-f="date" placeholder="09/05" value="${esc(p.date)}" size="6"></label>
            <label>卡號<input data-f="card" value="${esc(p.card)}" size="6"></label>
            <label>單位<input data-f="unit" value="${esc(p.unit)}"></label>
            <label class="check"><input type="checkbox" data-f="include" ${p.include ? 'checked' : ''}>納入計算</label>
            <div class="problem small" data-problem>${esc(pageProblem(p))}</div>
          </div>
        </div>`).join('')}`;
  }

  const pageList = $('ip-pdf-list');
  pageList.addEventListener('input', onPageEdit);
  pageList.addEventListener('change', onPageEdit);
  function onPageEdit(e) {
    const el = e.target.closest('[data-f]');
    const card = e.target.closest('.pdf-card');
    if (!el || !card) return;
    const p = state.pages.find((x) => x.id === Number(card.dataset.id));
    const f = el.dataset.f;
    p[f] = f === 'include' ? el.checked : el.value;

    // 輸入教師姓名 → 自動帶卡號、單位
    let hint = '';
    if (f === 'teacher' && state.parsed) {
      const hit = Calc.lookupCard(state.parsed.directory, p.teacher);
      if (hit.card) {
        p.card = hit.card; p.unit = hit.unit;
        card.querySelector('[data-f="card"]').value = p.card;
        card.querySelector('[data-f="unit"]').value = p.unit;
      } else if (hit.candidates.length > 1) {
        hint = `同名多人，請填卡號：${hit.candidates.map((c) => `${c.card}(${c.unit})`).join('、')}`;
      }
    }
    card.classList.toggle('excluded', !p.include);
    card.querySelector('[data-problem]').textContent = hint || pageProblem(p);
    recalc();
  }

  pageList.addEventListener('click', (e) => {
    const img = e.target.closest('img[data-full]');
    if (!img) return;
    const p = state.pages.find((x) => x.id === Number(img.dataset.full));
    const viewer = $('img-viewer');
    viewer.querySelector('img').src = p.fullUrl;
    viewer.hidden = false;
  });
  $('img-viewer').addEventListener('click', () => { $('img-viewer').hidden = true; });

  function renderResult() {
    $('ip-result-step').hidden = false;
    const s = state.summary;
    const total = s.reduce((a, x) => a + (x.amount || 0), 0);
    const hand = state.pages.filter((p) => p.include && p.teacher.trim() && parsePageDate(p.date)).length;
    const sk = state.parsed.skipped;
    $('ip-stats').innerHTML = `
      <div class="stat"><span>老師人數</span><b>${s.length}</b></div>
      <div class="stat"><span>總金額</span><b>${money(total)}</b></div>
      <div class="stat"><span>計費依據</span><b>${app.hasData() ? `申請單 ${app.matched().filter((m) => m.include !== false).length} 列` : `系統 ${money(state.parsed.records.length)} 筆`}</b></div>
      <div class="stat"><span>手寫紀錄（已填/總頁）</span><b>${hand} / ${state.pages.length}</b></div>
      <div class="stat muted-stat"><span>未評核完成（不計）</span><b>${money(sk.notDone)}</b></div>`;

    const warns = [];
    if (!app.hasData()) warns.push('尚未匯入申請單 PDF：目前暫以 rawdata 全部已評核紀錄計算（正式計費請匯入申請單）。');
    if (app.busy()) warns.push('申請單辨識中…結果會自動更新。');
    const noCardApp = s.filter((x) => !x.card && x.highlight).length;
    if (noCardApp) warns.push(`有 ${noCardApp} 列為申請單上無法辨識或找不到卡號的老師，下載的名冊中為<b>紅字、卡號空白</b>（日期無法辨識者金額也空白），請人工核對。`);
    if (state.unclaimed.length) {
      const tea = [...new Set(state.unclaimed.map((r) => r.name))];
      warns.push(`系統有紀錄、但申請單沒有列的：${state.unclaimed.length} 筆（${tea.length} 位老師，不計費）：${tea.slice(0, 15).map(esc).join('、')}${tea.length > 15 ? '…' : ''}`);
    }
    const red = s.filter((x) => x.highlight);
    if (red.length) warns.push(`紅字 ${red.length} 位：金額含「申請單有、系統沒有已評核紀錄」的教學日`);
    const noCard = s.filter((x) => !x.card);
    const badPages = state.pages.filter((p) => p.include && pageProblem(p) && !pageProblem(p).startsWith('缺卡號'));
    if (badPages.length) warns.push(`有 ${badPages.length} 頁手寫紀錄資料不完整，尚未算入`);
    if (!state.template) warns.push('未匯入範本：下載時會產生同格式的空白名冊（年、月、單位、卡號、姓名、金額、備註、教學日）。');
    else if (s.length > tinfo().capacity) {
      warns.push(`人數 ${s.length} 超過範本 ${tinfo().capacity} 列，下載時會自動增加列數並更新合計公式`);
    }
    $('ip-warnings').innerHTML = warns.map((w) => `<div class="warn">⚠ ${w}</div>`).join('');

    const q = $('ip-filter').value.trim();
    const rows = q ? s.filter((x) => `${x.unit} ${x.card} ${x.name}`.includes(q)) : s;
    // 與下載的名冊相同：年｜月｜單位｜卡號｜姓名｜金額｜備註｜教學日；紅字列同 Excel；滑鼠停留可看說明
    const ti = tinfo();
    $('ip-result').innerHTML = rows.length ? rows.map((x) => {
      const why = [x.highlight ? (x.amount === null ? '日期無法辨識' : !x.card ? '無法辨識／缺卡號' : `系統無紀錄 +${money(x.appExtra)}`) : '', `${x.weeks.length} 週`].filter(Boolean).join('，');
      return `<tr class="${x.highlight ? 'row-red' : ''}" title="${esc(why)}">
        <td>${ti.rocYear}</td><td>${ti.month}</td><td>${esc(x.unit)}</td><td>${esc(x.card)}</td><td>${esc(x.name)}</td>
        <td class="num">${x.amount === null ? '' : money(x.amount)}</td><td>匯入銀行帳號</td>
        <td class="days">${esc(x.days.length ? Calc.formatDays(x.days) : (x.daysText || ''))}</td></tr>`;
    }).join('') + `<tr class="total-row"><td colspan="4"></td><td>合計</td><td class="num">${money(total)}</td><td colspan="2"></td></tr>`
      : '<tr class="empty"><td colspan="8">沒有符合的資料</td></tr>';
  }

  $('ip-filter').addEventListener('input', () => state.parsed && renderResult());
  $('ip-week-start').addEventListener('change', recalc);
  $('ip-month').addEventListener('change', reparse);

  // ---------- 下載比對後的 rawdata ----------
  $('ip-raw-download').addEventListener('click', async () => {
    if (!state.raw || !app.hasData()) { alert('請先匯入 rawdata 與申請單 PDF'); return; }
    try {
      await AppPdf.downloadMatchedRaw(state.raw, app.matchedRowNums());
    } catch (err) { console.error(err); alert(`產生檔案失敗：${err.message}`); }
  });

  // ---------- 下載 ----------
  $('ip-download').addEventListener('click', async () => {
    if (!state.summary.length) return;
    const btn = $('ip-download');
    btn.disabled = true;
    try {
      // 有範本填範本；沒有範本產生同格式的空白名冊
      let wb;
      if (state.template) { wb = new ExcelJS.Workbook(); await wb.xlsx.load(state.template.buffer.slice(0)); } else {
        const ti = tinfo();
        wb = Fill.blankTemplate(ExcelJS, { title: '教學住診指導費支領名冊', rocYear: ti.rocYear, month: ti.month, withDays: true });
      }
      Fill.fillTemplate(wb, state.summary);
      const out = await wb.xlsx.writeBuffer();
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = state.template ? `${state.template.name.replace(/\.xlsx$/i, '')}_已填寫.xlsx` : `教學住診名冊_${tinfo().rocYear}${String(tinfo().month).padStart(2, '0')}.xlsx`;
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
  function reset() {
    state = emptyState();
    app.reset();
    $('ip-app-step').hidden = true;
    $('ip-filter').value = '';
    $('ip-result').innerHTML = '';
    $('img-viewer').hidden = true;
    reparse();
  }

  App.registerTab('teaching-clinic', { reset });
  renderStatus();
})();
