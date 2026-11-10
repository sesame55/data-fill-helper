/*
 * 分頁：導生座談
 * 流程：匯入（rawdata ×3 / 導談完成率統計 / 簽到簿 PDF）→ 輸入手寫簽到 → 結果 → 下載
 * 所有資料只存在記憶體（state），清空、關閉分頁即消失。
 */
(() => {
  const T = window.TutorCalc;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = (n) => n.toLocaleString('en-US');
  const mmdd = (iso) => (iso ? `${iso.slice(5, 7)}/${iso.slice(8, 10)}` : '');
  const SIGN_CATS = [['PGY', 'PGY'], ['R', '住院醫師'], ['UGY', '實習生'], ['臨培', '臨培（醫事人員）'], ['青果森', '青果森']];

  const emptyState = () => ({
    raws: {},          // cat → { name, records }
    roster: null,      // { name, parsed }
    pages: [],         // { key, file, sheetCat, page, imgUrl, entries: [{ id, line, date, card, name, illegible, note }] }
    result: null,
    dupChoice: {},     // 重複申請：卡號 → 'all' 或保留那筆的 signKey
    busy: [],
  });
  let state = emptyState();
  let seq = 0;

  // ---------- 匯入 ----------
  const dropzone = $('tu-dropzone');
  const input = $('tu-file');
  input.addEventListener('change', () => { handleFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
  dropzone.addEventListener('drop', (e) => handleFiles([...e.dataTransfer.files]));

  async function handleFiles(files) {
    for (const f of files) {
      const ext = f.name.split('.').pop().toLowerCase();
      setBusy(f.name, true);
      try {
        if (ext === 'pdf') await loadPdf(f);
        else if (ext === 'xlsx') await loadXlsx(f);
        else alert(`不支援的檔案：${f.name}`);
      } catch (err) {
        console.error(err);
        alert(`讀取「${f.name}」失敗：${err.message}`);
      } finally {
        setBusy(f.name, false);
      }
    }
    guessMonth();
    renderStatus();
    renderPages();
    recompute();
    runOcr();
  }

  async function loadXlsx(file) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { cellDates: true });
    const first = wb.Sheets[wb.SheetNames[0]];
    const type = T.detectRawType(XLSX.utils.sheet_to_json(first, { header: 1 })[0] || []);
    if (type) {
      state.raws[type] = { name: file.name, records: T.parseRaw(XLSX.utils.sheet_to_json(first, { defval: '', raw: false }), type) };
      return;
    }
    const sheets = wb.SheetNames.map((name) => ({ name, rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' }) }));
    const parsed = T.parseRoster(sheets);
    if (!parsed.people.size) throw new Error('看不出是 rawdata 或導談完成率統計（找不到 R檢核／PGY／C／醫事人員 工作表）');
    state.roster = { name: file.name, parsed, buffer };
  }

  async function loadPdf(file) {
    const sheetCat = T.signinSheetCat(file.name);
    state.pages = state.pages.filter((p) => p.file !== file.name);
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false }).promise;
    for (let p = 1; p <= pdf.numPages; p += 1) {
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale: 1.6 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport }).promise;
      const { entries, note, struck } = detectEntries(canvas, ctx);
      state.pages.push({
        key: ++seq, file: file.name, sheetCat, page: p,
        imgUrl: canvas.toDataURL('image/jpeg', 0.85),
        entries, note, struck: struck || 0,
      });
      canvas.width = 0;
    }
    pdf.destroy();
  }

  const newEntry = (line) => ({ id: ++seq, line, date: '', card: '', name: '', illegible: false, note: '', touched: {} });

  // ---------- 自動偵測：找出有手寫的列、截圖、排入文字辨識 ----------
  const ocrQueue = [];   // { entryId, field, canvas }
  function crop(src, x, y, w, h, scale = 1) {
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(w * scale)); c.height = Math.max(1, Math.round(h * scale));
    const cx = c.getContext('2d');
    cx.fillStyle = '#fff'; cx.fillRect(0, 0, c.width, c.height);
    cx.drawImage(src, x, y, w, h, 0, 0, c.width, c.height);
    return c;
  }
  /** 辨識前處理：轉黑白、四周加白邊 */
  function forOcr(src) {
    const padPx = 16;
    const c = document.createElement('canvas');
    c.width = src.width + padPx * 2; c.height = src.height + padPx * 2;
    const cx = c.getContext('2d');
    cx.fillStyle = '#fff'; cx.fillRect(0, 0, c.width, c.height);
    cx.drawImage(src, padPx, padPx);   // 只加白邊；轉黑白、分離藍筆實測都會降低辨識率
    return c;
  }

  function detectEntries(canvas, ctx) {
    if (!window.SheetDetect) return { entries: [newEntry(1)], note: '' };
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const det = SheetDetect.detect(img);
    if (!det) return { entries: [newEntry(1)], note: '此頁沒有偵測到簽到表格（可能是請款總表）；如有簽到請手動新增。' };
    const cr = SheetDetect.columnRects(det);
    // 整列刪除線＝刪除此筆申請，不列出（只計數）
    let struck = 0; const struckLines = [];
    const rowsKept = SheetDetect.filledRows(det).filter((r) => {
      const hit = SheetDetect.isStruck(img, det.cols[1], det.cols[det.cols.length - 1], r.top, r.bottom);
      if (hit) { struck += 1; struckLines.push(r.index); }
      return !hit;
    });
    const entries = rowsKept.map((r) => {
      const e = newEntry(r.index);
      e.auto = true;
      e.stripUrl = crop(canvas, det.left, r.top, det.right - det.left, r.bottom - r.top).toDataURL('image/jpeg', 0.85);
      if (cr) {
        const m = 5;   // 內縮避開格線
        e.ocrPending = 5; e.ocr = {};
        [['date', cr.date], ['name', cr.name], ['card', cr.card], ['gsm', cr.gsm], ['tutor', cr.tutor]].forEach(([field, c]) => {
          ocrQueue.push({ entryId: e.id, field, canvas: crop(canvas, c.x0 + 3, r.top + 3, c.x1 - c.x0 - 6, r.bottom - r.top - 4, 1.25) });
        });
      }
      return e;
    });
    const note = det.rows.length !== 21 ? '表格線偵測可能不完整，請對照左側整頁影像確認有無漏列。' : '';
    return { entries: entries.length ? entries : [newEntry(1)], note: [note, struck ? `刪除線：第 ${struckLines.join('、')} 列` : ''].filter(Boolean).join('；'), struck };
  }

  // 文字辨識（Tesseract，於本機瀏覽器執行，模型檔在 lib/tesseract）
  let ocrWorker = null; let ocrRunning = false;
  const lev = (a, b) => {
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
    for (let j = 1; j <= b.length; j += 1) d[0][j] = j;
    for (let i = 1; i <= a.length; i += 1) for (let j = 1; j <= b.length; j += 1) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    return d[a.length][b.length];
  };
  /** OCR 卡號 → 名冊／rawdata 中最接近的卡號（距離 ≤1 且唯一才採用） */
  function matchCard(raw, sheetCat) {
    const s = raw.toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/O/g, '0').replace(/[IL]/g, '1');
    if (s.length < 3) return { card: '', hint: raw };
    const want = { PGY: 'PGY', R: 'R', UGY: 'UGY', 臨培: '臨培' }[sheetCat];
    let best = []; let bestD = 9;
    directory.forEach((p, card) => {
      if (want && p.cat !== want) return;
      const dd = lev(s, card);
      if (dd < bestD) { bestD = dd; best = [card]; } else if (dd === bestD) best.push(card);
    });
    if (bestD <= 1 && best.length === 1) return { card: best[0], hint: '', cands: [] };
    // 候選：同類別中與辨識片段最接近的 3 位（片段可能只有前幾碼，所以也比對同長度前綴）
    const scored = [];
    directory.forEach((p, card) => {
      if (want && p.cat !== want) return;
      const d = Math.min(lev(s, card), lev(s, card.slice(0, s.length)) + 0.5 * Math.max(0, card.length - s.length));
      scored.push({ card, d });
    });
    scored.sort((a, b) => a.d - b.d);
    const cands = scored.filter((x) => x.d <= 2).slice(0, 3).map((x) => x.card);
    return { card: '', hint: s, cands };
  }
  /** 本月 rawdata 座談日期（卡號已知時當作日期建議） */
  function rawDateFor(card) {
    const ym = $('tu-month').value;
    const hit = rawAll().filter((r) => r.card === card && r.date.startsWith(ym)).sort((a, b) => a.date.localeCompare(b.date))[0];
    return hit ? `${hit.date.slice(5, 7)}/${hit.date.slice(8, 10)}` : '';
  }
  function applyCard(entry, card) {
    entry.card = card;
    entry.name = directory.get(card)?.name || '';
    entry.cands = [];
    if (!entry.date && !entry.touched.date) entry.date = rawDateFor(card);
  }
  function parseOcrDate(raw) {
    const m = raw.replace(/\s/g, '').match(/(\d{1,2})[/\\|1lI]?(\d{1,2})/);
    const ym = $('tu-month').value;
    if (!m || !ym) return '';
    const mm = Number(m[1]); const dd = Number(m[2]);
    if (mm !== Number(ym.slice(5, 7)) || dd < 1 || dd > 31) return '';
    return `${String(mm).padStart(2, '0')}/${String(dd).padStart(2, '0')}`;
  }
  /**
   * 手寫簽到列的綜合比對：卡號（字母＋數字）、GSM（5 碼數字）、日期三欄交叉比對名冊與 rawdata。
   * 單一欄位手寫辨識常出錯，但三欄一起看，正確的人通常會明顯領先。
   */
  const digitSim = (a, b) => (a && b ? 1 - lev(a, b) / Math.max(a.length, b.length) : 0);
  function resolveEntry(page, entry) {
    const o = entry.ocr || {};
    // 手寫常見混淆：( → 1、b → 6、s → 5、q → 9、Z → 2、o → 0；卡號首字 V → U
    const fixDigits = (t) => String(t || '').normalize('NFKC').replace(/[(（|lI!]/g, '1').replace(/[oO。]/g, '0').replace(/[bG]/g, '6').replace(/[sS]/g, '5').replace(/q/g, '9').replace(/[zZ]/g, '2');
    o.date = fixDigits(o.date);
    const date = parseOcrDate(o.date || '');
    if (date && !entry.touched.date) entry.date = date;
    const c0 = String(o.card || '').normalize('NFKC').trim();
    const cardTxt = (/^[VvUu]/.test(c0) ? `U${fixDigits(c0.slice(1))}` : /^[A-Za-z]/.test(c0) ? c0[0].toUpperCase() + fixDigits(c0.slice(1)) : fixDigits(c0)).toUpperCase().replace(/[^A-Z0-9]/g, '');
    const gsmTxt = fixDigits(o.gsm).replace(/\D/g, '');
    const nameTxt = String(o.name || '').normalize('NFKC').replace(/[^\p{Script=Han}]/gu, '');
    const tutorTxt = String(o.tutor || '').normalize('NFKC').replace(/[^\p{Script=Han}]/gu, '');
    const first = matchCard(cardTxt || '', page.sheetCat);
    if (first.card) { applyCard(entry, first.card); return; }
    const want = { PGY: 'PGY', R: 'R', UGY: 'UGY', 臨培: '臨培' }[page.sheetCat];
    const ym = $('tu-month').value;
    const rawDates = new Map();
    rawAll().forEach((r) => { if (r.date.startsWith(ym)) (rawDates.get(r.card) || rawDates.set(r.card, new Set()).get(r.card)).add(`${r.date.slice(5, 7)}/${r.date.slice(8, 10)}`); });
    const scored = [];
    directory.forEach((p, card) => {
      if (want && p.cat !== want) return;
      let sc = 0;
      if (cardTxt.length >= 2) sc += Math.max(digitSim(cardTxt, card), digitSim(cardTxt, card.slice(0, cardTxt.length)) * 0.8);
      if (gsmTxt.length >= 3 && p.gsm) sc += digitSim(gsmTxt, p.gsm);
      if (nameTxt.length >= 2 && p.name) sc += AppForm.sim(nameTxt, p.name);   // 姓名（字重疊比例）
      if (tutorTxt.length >= 2 && p.tutor) sc += 0.8 * AppForm.sim(tutorTxt, p.tutor);   // 輔導老師（導師）
      const ds = rawDates.get(card);
      if (ds) sc += (date && ds.has(date)) ? 0.8 : 0.3;   // 當月有座談紀錄；日期相同更可信
      scored.push({ card, sc });
    });
    scored.sort((a, b) => b.sc - a.sc);
    const [b1, b2] = scored;
    if (b1 && b1.sc >= 1.2 && b1.sc - (b2?.sc || 0) >= 0.3) { applyCard(entry, b1.card); entry.autoJoint = true; return; }
    entry.ocrHint = [nameTxt && `姓名 ${nameTxt}`, tutorTxt && `導師 ${tutorTxt}`, cardTxt && `卡號 ${cardTxt}`, gsmTxt && `GSM ${gsmTxt}`, o.date && `日期 ${o.date}`].filter(Boolean).join('、');
    entry.cands = scored.slice(0, 3).filter((x) => x.sc >= 0.6).map((x) => x.card);
  }

  async function runOcr() {
    if (ocrRunning || !ocrQueue.length || !window.PaddleRec) return;
    ocrRunning = true;
    try {
      setOcrStatus('載入手寫辨識模型（PaddleOCR，第一次約需數秒）…');
      await PaddleRec.init();
      directory = T.buildDirectory(roster(), rawAll());
      const total = ocrQueue.length; let done = 0;
      while (ocrQueue.length) {
        const job = ocrQueue.shift();
        done += 1;
        setOcrStatus(`文字辨識中… ${done}/${total}`);
        const page = state.pages.find((p) => p.entries.some((e) => e.id === job.entryId));
        const entry = page?.entries.find((e) => e.id === job.entryId);
        if (!entry || entry.touched[job.field]) continue;
        entry.ocr[job.field] = (await PaddleRec.recognize(job.canvas)).text.trim();
        entry.ocrPending -= 1;
        if (entry.ocrPending > 0) continue;
        resolveEntry(page, entry);   // 日期＋卡號＋GSM 三欄都辨識完 → 綜合比對
        updateEntryDom(page, entry);
        recomputeSoon();
      }
      renderPages();
      setOcrStatus('文字辨識完成：辨識不出卡號的列已列入「無法辨識」（下載的結果中為紅字），請人工審核。');
    } catch (err) {
      console.error(err);
      setOcrStatus(`文字辨識無法使用（${err.message}），請手動輸入。`);
    } finally {
      ocrRunning = false;
    }
  }
  function setOcrStatus(msg) { const el = $('tu-ocr-status'); if (el) el.textContent = msg; }
  function updateEntryDom(page, entry) {
    const tr = pagesEl.querySelector(`[data-entry="${entry.id}"]`);
    if (!tr) return;
    ['date', 'card', 'name'].forEach((f) => {
      const inp = tr.querySelector(`[data-f="${f}"]`);
      if (inp && document.activeElement !== inp && !entry.touched[f]) {
        inp.value = entry[f];
        inp.classList.toggle('auto-filled', !!entry[f]);
      }
    });
    tr.querySelector('[data-info]').innerHTML = infoHtml(entry);
  }
  /** 「自動帶出」欄：查到的人；或辨識片段＋候選按鈕 */
  function infoHtml(e) {
    const lk = lookupText(e.card);
    if (lk.info && e.card) return lk.info;
    const cands = (e.cands || []).map((c) => `<button class="cand" data-cand="${esc(c)}" title="點選帶入">${esc(c)} ${esc(directory.get(c)?.name || '')}</button>`).join('');
    const hint = e.ocrHint ? `<span class="muted">辨識：${esc(e.ocrHint)}？</span>` : '';
    return `${lk.info}${hint}${cands ? `<div>${cands}</div>` : ''}`;
  }

  function setBusy(name, on) {
    state.busy = on ? [...state.busy, name] : state.busy.filter((n) => n !== name);
    renderStatus();
  }

  function guessMonth() {
    const m = $('tu-month');
    if (m.value) return;
    const fromPdf = state.pages.map((p) => p.file.match(/(20\d{2})(\d{2})/)).find(Boolean);
    if (fromPdf) { m.value = `${fromPdf[1]}-${fromPdf[2]}`; return; }
    const dates = Object.values(state.raws).flatMap((r) => r.records.map((x) => x.date)).sort();
    if (dates.length) m.value = dates[dates.length - 1].slice(0, 7);
  }

  // ---------- 計算 ----------
  const rawAll = () => Object.values(state.raws).flatMap((r) => r.records);
  const roster = () => state.roster?.parsed || { people: new Map(), sheetsUsed: {} };
  let directory = new Map();

  function signins() {
    return state.pages.flatMap((p) => p.entries
      .filter((e) => e.card.trim() || e.name.trim() || e.illegible || e.stripUrl)
      .map((e) => {
        const failed = !e.card.trim();   // 自動辨識不出卡號 → 列入「無法辨識」（紅字），人工審核
        const hint = [e.ocrHint ? `辨識：${e.ocrHint}` : '', e.cands?.length ? `可能是 ${e.cands.map((c) => `${c} ${directory.get(c)?.name || ''}`).join('／')}` : ''].filter(Boolean).join('；');
        return { file: p.file, sheetCat: p.sheetCat, page: p.page, line: Number(e.line) || 0,
          date: e.date, card: e.card, name: e.name, illegible: e.illegible || failed, auto: !!e.autoJoint,
          note: e.note || (failed ? `自動辨識失敗${hint ? `（${hint}）` : ''}` : '') };
      }));
  }


  function recompute() {
    const ym = $('tu-month').value;
    directory = T.buildDirectory(roster(), rawAll());
    const ready = ym && (Object.keys(state.raws).length || state.pages.length);
    state.result = ready ? T.compute({ ym, rawAll: rawAll(), roster: roster(), signins: signins(), dupChoice: state.dupChoice }) : null;
    renderResult();
  }
  let timer = null;
  const recomputeSoon = () => { clearTimeout(timer); timer = setTimeout(recompute, 250); };

  // ---------- 畫面：狀態 ----------
  function renderStatus() {
    const item = (ok, label, detail) =>
      `<div class="status-item ${ok ? 'ok' : ''}"><span class="dot"></span><b>${label}</b><span class="muted small">${detail}</span></div>`;
    const raw = (cat, label) => item(!!state.raws[cat], `rawdata ${label}`,
      state.raws[cat] ? `${esc(state.raws[cat].name)}・${money(state.raws[cat].records.length)} 筆` : '尚未匯入');
    const r = state.roster;
    const files = [...new Set(state.pages.map((p) => p.file))];
    $('tu-status').innerHTML = [
      raw('PGY', 'PGY'), raw('R', '住院醫師'), raw('UGY', '實習生'),
      item(!!r, '名冊', r ? `${esc(r.name)}・${money(r.parsed.people.size)} 人（${Object.entries(r.parsed.sheetsUsed).map(([c, s]) => `${T.CAT_LABEL[c]}：${esc(s)}`).join('、')}）` : '尚未匯入（建議匯入，才能列出全部學員與名冊列號）'),
      item(files.length > 0, '簽到簿', files.length ? `${files.length} 個檔案・${state.pages.length} 頁` : '尚未匯入'),
      state.busy.length ? `<div class="status-item busy"><span class="dot"></span>讀取中：${esc(state.busy.join('、'))}</div>` : '',
    ].join('');
  }

  // ---------- 畫面：簽到輸入 ----------
  function lookupText(card) {
    const c = T.normCard(card);
    if (!c) return { name: '', info: '' };
    const p = directory.get(c);
    if (!p) return { name: '', info: '<span class="problem">查無此卡號</span>' };
    return { name: p.name, info: `${esc(T.CAT_LABEL[p.cat] || p.cat)}・導師 ${esc(p.tutor || '—')}` };
  }

  function entryRow(e) {
    const lk = lookupText(e.card);
    const auto = (f) => (e.auto && e[f] && !e.touched[f] ? 'auto-filled' : '');
    const strip = e.stripUrl ? `<tr class="strip-row" data-strip="${e.id}"><td colspan="8"><img src="${e.stripUrl}" alt="第 ${esc(e.line)} 列手寫影像"></td></tr>` : '';
    return `${strip}<tr data-entry="${e.id}" class="${e.illegible ? 'row-warn' : ''}">
      <td><input data-f="line" value="${esc(e.line)}" size="2"></td>
      <td><input data-f="date" class="${auto('date')}" value="${esc(e.date)}" placeholder="10/20" size="5"></td>
      <td><input data-f="card" class="${auto('card')}" value="${esc(e.card)}" placeholder="卡號" size="6"></td>
      <td><input data-f="name" class="${auto('name')}" value="${esc(e.name || lk.name)}" placeholder="姓名" size="6"></td>
      <td class="small" data-info>${infoHtml(e)}</td>
      <td><input type="checkbox" data-f="illegible" ${e.illegible ? 'checked' : ''} title="無法辨識"></td>
      <td><input data-f="note" value="${esc(e.note)}" placeholder="說明" size="10"></td>
      <td><button class="btn-ghost small" data-del title="刪除此列">✕</button></td>
    </tr>`;
  }

  function renderPages() {
    $('tu-sign-step').hidden = state.pages.length === 0;
    const rows = state.pages.map((p) => {
      const auto = p.entries.filter((e) => e.stripUrl);
      const ok = auto.filter((e) => e.card.trim()).length;
      return `<tr><td>${esc(p.file)}</td><td>${p.page}</td><td>${esc(SIGN_CATS.find(([v]) => v === p.sheetCat)?.[1] || '（未知）')}</td>
        <td class="num">${auto.length}</td><td class="num">${ok}</td><td class="num ${auto.length - ok ? 'text-red' : ''}">${auto.length - ok}</td><td class="num">${p.struck || ''}</td>
        <td class="small muted">${esc(p.note || '')}</td></tr>`;
    }).join('');
    $('tu-pages').innerHTML = `<div class="table-wrap"><table class="file-table">
      <thead><tr><th>簽到簿</th><th>頁</th><th>類別</th><th>偵測到的簽到列</th><th>自動辨識</th><th>無法辨識（紅字）</th><th>刪除線（不列出）</th><th>備註</th></tr></thead>
      <tbody>${rows}</tbody></table></div>`;
  }


  const pagesEl = $('tu-pages');
  const findPage = (el) => state.pages.find((p) => p.key === Number(el.closest('[data-page]').dataset.page));

  pagesEl.addEventListener('input', (e) => {
    const f = e.target.dataset.f;
    const tr = e.target.closest('[data-entry]');
    if (!f || !tr) return;
    const page = findPage(tr);
    const entry = page.entries.find((x) => x.id === Number(tr.dataset.entry));
    entry[f] = f === 'illegible' ? e.target.checked : e.target.value;
    entry.touched[f] = true;
    e.target.classList.remove('auto-filled');
    if (f === 'card') {
      directory = T.buildDirectory(roster(), rawAll());
      entry.cands = [];
      const lk = lookupText(entry.card);
      tr.querySelector('[data-info]').innerHTML = infoHtml(entry);
      if (lk.name) {
        entry.name = lk.name; tr.querySelector('[data-f="name"]').value = lk.name;
        if (!entry.date && !entry.touched.date) {
          entry.date = rawDateFor(T.normCard(entry.card));
          const di = tr.querySelector('[data-f="date"]'); di.value = entry.date; di.classList.toggle('auto-filled', !!entry.date);
        }
      }
    }
    if (f === 'illegible') tr.classList.toggle('row-warn', entry.illegible);
    recomputeSoon();
  });
  pagesEl.addEventListener('change', (e) => {
    if (e.target.matches('[data-f="illegible"]')) e.target.dispatchEvent(new Event('input', { bubbles: true }));
    if (e.target.matches('[data-cat]')) { findPage(e.target).sheetCat = e.target.value; recomputeSoon(); }
  });
  pagesEl.addEventListener('click', (e) => {
    if (e.target.matches('[data-zoom]')) {
      const v = $('img-viewer'); v.querySelector('img').src = e.target.src; v.hidden = false; return;
    }
    if (e.target.matches('[data-add]')) { addRow(findPage(e.target)); return; }
    const candBtn = e.target.closest('[data-cand]');
    if (candBtn) {
      const page = findPage(candBtn);
      const entry = page.entries.find((x) => x.id === Number(candBtn.closest('[data-entry]').dataset.entry));
      applyCard(entry, candBtn.dataset.cand);
      entry.touched.card = true; entry.touched.name = true;
      const tr = candBtn.closest('[data-entry]');
      tr.querySelector('[data-f="card"]').value = entry.card;
      tr.querySelector('[data-f="name"]').value = entry.name;
      const di = tr.querySelector('[data-f="date"]');
      if (!entry.touched.date) { di.value = entry.date; di.classList.toggle('auto-filled', !!entry.date); }
      tr.querySelector('[data-info]').innerHTML = infoHtml(entry);
      recomputeSoon();
      return;
    }
    if (e.target.matches('[data-del]')) {
      const page = findPage(e.target);
      const id = Number(e.target.closest('[data-entry]').dataset.entry);
      page.entries = page.entries.filter((x) => x.id !== id);
      if (!page.entries.length) page.entries.push(newEntry(1));
      rerenderPage(page);
      recomputeSoon();
    }
  });
  pagesEl.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter' || !e.target.matches('input[data-f]')) return;
    e.preventDefault();
    const tr = e.target.closest('[data-entry]');
    const page = findPage(tr);
    const rows = [...tr.parentElement.querySelectorAll('tr[data-entry]')];
    const next = rows[rows.indexOf(tr) + 1];
    if (!next) addRow(page);
    else next.querySelector('[data-f="date"]').focus();
  });

  function addRow(page) {
    const last = page.entries[page.entries.length - 1];
    const e = newEntry((Number(last?.line) || 0) + 1);
    if (last?.date) e.date = last.date;      // 同一頁常是同一天，先帶入上一列日期
    page.entries.push(e);
    rerenderPage(page);
    pagesEl.querySelector(`[data-entry="${e.id}"] [data-f="card"]`).focus();
  }
  function rerenderPage(page) {
    pagesEl.querySelector(`[data-page="${page.key}"] tbody`).innerHTML = page.entries.map(entryRow).join('');
  }

  // ---------- 畫面：結果 ----------
  function renderResult() {
    const r = state.result;
    $('tu-result-step').hidden = !r;
    if (!r) return;
    const count = (s) => r.rows.filter((x) => x.status === s).length;
    const red = r.rows.filter((x) => x.red);
    const total = r.meal.reduce((a, x) => a + x.amountAfter, 0);
    $('tu-stats').innerHTML = `
      <div class="stat"><span>V（系統有紀錄）</span><b>${count('V')}</b></div>
      <div class="stat"><span>院內（有簽到）</span><b>${count('院內')}</b></div>
      <div class="stat"><span>簽到但系統無紀錄</span><b class="text-red">${red.length}</b></div>
      <div class="stat"><span>重複申請（人／扣除筆數）</span><b>${r.dupGroups.length} / ${r.duplicates.length}</b></div>
      <div class="stat"><span>無法辨識</span><b>${r.illegible.length}</b></div>
      <div class="stat"><span>餐費（修正後）</span><b>${money(total)}</b></div>`;

    const warns = [];
    if (!state.roster) warns.push('尚未匯入導談完成率統計（名冊），檢核結果只會列出本月有紀錄或有簽到的人。');
    ['PGY', 'R', 'UGY'].filter((c) => !state.raws[c]).forEach((c) => warns.push(`尚未匯入 ${T.CAT_LABEL[c]} rawdata，該類學員無法判斷 V。`));
    state.pages.filter((p) => !p.sheetCat).forEach((p) => warns.push(`${esc(p.file)} 無法從檔名判斷類別，請在簽到簿上方選擇。`));
    if (r.unknownGroup.length) warns.push(`有 ${r.unknownGroup.length} 筆簽到無法判斷餐費類別：${r.unknownGroup.map((s) => esc(s.card)).join('、')}`);
    const pending = state.pages.flatMap((p) => p.entries.filter((e) => e.stripUrl && !e.card.trim() && !e.illegible).map((e) => `${esc(p.file)} 第${p.page}頁第${e.line}列`));
    if (pending.length) warns.push(`有 ${pending.length} 列偵測到手寫但還沒有卡號（未計入）：${pending.slice(0, 12).join('、')}${pending.length > 12 ? '…' : ''}`);
    const notFound = signins().filter((s) => !s.illegible && s.card.trim() && !directory.get(T.normCard(s.card)));
    if (notFound.length) warns.push(`有 ${notFound.length} 筆簽到的卡號在名冊與 rawdata 都查不到：${notFound.map((s) => `${esc(s.card)}（${esc(s.file)} 第${s.page}頁第${s.line}列）`).join('、')}`);
    $('tu-warnings').innerHTML = warns.map((w) => `<div class="warn">⚠ ${w}</div>`).join('');

    $('tu-meal').innerHTML = r.meal.map((x) => `<tr>
      <td>${x.group}</td><td class="num">${x.before}</td><td class="num ${x.before !== x.after ? 'text-red' : ''}">${x.after}</td>
      <td class="num">${x.rate}</td><td class="num">${money(x.amountBefore)}</td><td class="num">${money(x.amountAfter)}</td></tr>`).join('')
      + `<tr><td><b>總和</b></td><td class="num">${r.meal.reduce((a, x) => a + x.before, 0)}</td><td class="num">${r.meal.reduce((a, x) => a + x.after, 0)}</td><td></td>
        <td class="num">${money(r.meal.reduce((a, x) => a + x.amountBefore, 0))}</td><td class="num"><b>${money(total)}</b></td></tr>`;
    $('tu-message').value = T.mealMessage(r);

    const list = (title, items) => (items.length ? `<h4>${title}</h4><ul class="plain-list">${items.join('')}</ul>` : '');
    $('tu-lists').innerHTML = [
      dupTable(r.dupGroups),
      // 與下載檔「無法辨識」工作表相同欄位，紅字
      r.illegible.length ? `<h4>無法辨識（${r.illegible.length}）</h4><div class="table-wrap"><table class="file-table"><thead><tr><th>檔案</th><th>頁</th><th>列</th><th>日期</th><th>卡號</th><th>姓名</th><th>說明</th></tr></thead><tbody>${r.illegible.map((s) => `<tr class="row-red"><td>${esc(s.file)}</td><td>${s.page}</td><td>${s.line}</td><td>${esc(s.date)}</td><td>${esc(s.card)}</td><td>${esc(s.name)}</td><td>${esc(s.note || '無法辨識')}</td></tr>`).join('')}</tbody></table></div>` : '',
    ].join('');

    const q = $('tu-filter').value.trim();
    const only = $('tu-only-status').checked;
    const rows = r.rows.filter((x) => (!only || x.status) && (!q || `${T.CAT_LABEL[x.cat]} ${x.unit} ${x.card} ${x.name} ${x.tutor}`.includes(q)));
    // 與下載檔「檢核結果」工作表相同欄位
    $('tu-result').innerHTML = rows.length ? rows.slice(0, 1000).map((x) => `<tr>
      <td>${esc(T.CAT_LABEL[x.cat] || x.cat)}</td><td class="small">${esc(x.sheet || '')}</td><td class="small">${x.row ?? ''}</td><td>${esc(x.unit)}</td><td>${esc(x.card)}</td>
      <td>${esc(x.name)}</td><td>${esc(x.tutor)}</td>
      <td>${x.red ? `<span class="red-tag">${x.status}</span>` : esc(x.status)}</td><td>${x.date ? x.date.replace(/-/g, '/') : ''}</td>
      <td class="num">${x.rawCount || ''}</td><td class="small">${esc(x.signSources)}</td><td class="small">${esc(x.notes.join('；'))}</td></tr>`).join('')
      : '<tr class="empty"><td colspan="12">沒有符合的資料</td></tr>';
  }

  /** 重複申請：每組並列所有簽到，讓使用者選擇保留哪一筆或全部都算 */
  function dupTable(groups) {
    if (!groups.length) return '';
    return `<h4>重複申請（請選擇保留哪一筆；預設保留美德耐）</h4>
      <div class="table-wrap"><table class="file-table dup-table">
        <thead><tr><th>卡號</th><th>姓名</th><th>類別</th><th>簽到簿</th><th>頁／列</th><th>日期</th><th>保留</th></tr></thead>
        <tbody>${groups.map((g) => g.entries.map((s, i) => {
          const key = T.signKey(s);
          const dropped = g.keep !== 'all' && g.keep !== key;
          return `<tr class="${i === 0 ? 'group-start' : ''} ${dropped ? 'row-dropped' : ''}">
            ${i === 0 ? `<td rowspan="${g.entries.length + 1}">${esc(g.card)}</td><td rowspan="${g.entries.length + 1}">${esc(g.name)}</td><td rowspan="${g.entries.length + 1}">${esc(T.CAT_LABEL[g.cat] || g.cat)}</td>` : ''}
            <td>${esc(s.file)}</td><td>${s.page}／${s.line}</td><td>${esc(s.date)}</td>
            <td><label class="check-inline"><input type="radio" name="dup-${esc(g.card)}" data-dup="${esc(g.card)}" value="${esc(key)}" ${g.keep === key ? 'checked' : ''}>保留${dropped ? '<span class="text-red small">（扣除）</span>' : ''}</label></td>
          </tr>`;
        }).join('') + `<tr><td colspan="4" class="small muted">與上列為同一人</td>
            <td><label class="check-inline"><input type="radio" name="dup-${esc(g.card)}" data-dup="${esc(g.card)}" value="all" ${g.keep === 'all' ? 'checked' : ''}>全部都算</label></td></tr>`).join('')}
        </tbody></table></div>`;
  }
  $('tu-lists').addEventListener('change', (e) => {
    const r = e.target.closest('[data-dup]');
    if (!r) return;
    state.dupChoice[r.dataset.dup] = r.value;
    recompute();
  });

  $('tu-filter').addEventListener('input', renderResult);
  $('tu-only-status').addEventListener('change', renderResult);
  $('tu-month').addEventListener('change', recompute);
  $('tu-copy').addEventListener('click', async () => {
    try { await navigator.clipboard.writeText($('tu-message').value); } catch { $('tu-message').select(); document.execCommand('copy'); }
  });

  // ---------- 下載 ----------
  $('tu-download').addEventListener('click', async () => {
    const r = state.result;
    if (!r) return;
    const btn = $('tu-download');
    btn.disabled = true;
    try {
      const out = await T.buildWorkbook(ExcelJS, r).xlsx.writeBuffer();
      const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = `導生座談檢核_${Number(r.ym.slice(0, 4)) - 1911}年${r.ym.slice(5, 7)}月.xlsx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error(err);
      alert(`產生檔案失敗：${err.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // ---------- 下載：在原始統計檔填入「院內」 ----------
  $('tu-fill-download').addEventListener('click', async () => {
    const r = state.result;
    if (!r) return;
    if (!state.roster) { alert('請先匯入導談完成率統計檔（可用①下載的已填入檔案）'); return; }
    const btn = $('tu-fill-download');
    btn.disabled = true;
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(state.roster.buffer.slice(0));
      const rep = window.TutorFill.planSignFill(wb, r.rows, r.ym);
      const out = await window.TutorFill.patchXlsx(state.roster.buffer, rep.writes, { JSZip: window.JSZip, DOMParser, XMLSerializer });
      const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = `${state.roster.name.replace(/\.xlsx$/i, '')}_${Number(r.ym.slice(0, 4)) - 1911}${r.ym.slice(5, 7)}院內.xlsx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      const li = (x) => `<li>${esc(x.sheet)} 第${x.row}列 ${esc(x.card)} ${esc(x.name)}${x.existing ? `：表上為「${esc(x.existing)}」` : ''}</li>`;
      $('tu-fill-report').innerHTML = `<div class="info-box small">
        已寫入「院內」${rep.written.length} 人（其中紅底白字 ${rep.written.filter((x) => x.red).length} 人；原為 V 改成院內 ${rep.written.filter((x) => x.before === 'V').length} 人）。
        ${rep.missing.map((m) => `<div class="problem">⚠ ${esc(m)}</div>`).join('')}
        ${rep.kept.length ? `<div>未覆寫（表上已有其他內容）：</div><ul class="plain-list">${rep.kept.map(li).join('')}</ul>` : ''}
        ${rep.notInRoster.length ? `<div>名冊找不到：</div><ul class="plain-list">${rep.notInRoster.map((x) => `<li>${esc(x.card)} ${esc(x.name)}</li>`).join('')}</ul>` : ''}
      </div>`;
    } catch (err) {
      console.error(err);
      alert(`產生檔案失敗：${err.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // ---------- 清空 ----------
  App.registerTab('tutor', {
    reset() {
      state = emptyState();
      ocrQueue.length = 0;
      setOcrStatus('');
      $('tu-month').value = '';
      $('tu-filter').value = '';
      renderStatus(); renderPages(); recompute();
    },
  });
  renderStatus();
})();
