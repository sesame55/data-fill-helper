/*
 * 補助費申請單（PDF 掃描）：辨識＋比對（教學住診、教學門診共用）
 * AppPdf.create({ root, getRawRows, getYm, onChange })
 * - 自動轉正、偵測表格、文字辨識（lib/tesseract，本機執行不外傳）、與 rawdata 比對。
 * - 辨識方式（實測最準）：先清除表格水平線（文字常壓在格線上），再「逐格」辨識：
 *     日期格用數字模型、姓名格用中文模型（單行）。
 * - 兩種版面：系統印出版（日期｜主治醫師｜學員…）、Word 版（主治醫師｜學員（多位）｜看診日期 民國）。
 * - 畫面只顯示進度與統計；無法辨識的列直接計入，於下載的名冊中紅字／空白，供人工核對。
 */
(() => {
  const A = window.AppForm;
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

  // ---------- 文字辨識工作者（兩個分頁共用；兩組平行處理） ----------
  let pool = null;
  async function getPool(onStatus) {
    if (pool) return pool;
    const opt = { workerPath: 'lib/tesseract/worker.min.js', corePath: 'lib/tesseract', langPath: 'lib/tesseract/lang', gzip: true, workerBlobURL: false, cacheMethod: 'none' };
    onStatus('載入文字辨識模型（第一次約需數秒）…');
    const make = async (lang, params) => { const w = await Tesseract.createWorker(lang, 1, opt); await w.setParameters(params); return w; };
    pool = await Promise.all([0, 1].map(async () => ({
      num: await make('eng', { tessedit_pageseg_mode: '7', tessedit_char_whitelist: '0123456789-./' }),
      chi7: await make('chi_tra', { tessedit_pageseg_mode: '7' }),
      chi6: await make('chi_tra', { tessedit_pageseg_mode: '6' }),
      lat: await make('eng', { tessedit_pageseg_mode: '7', tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' }),
    })));
    return pool;
  }

  const rect = (x0, y0, x1, y1, m = 0) => ({ left: Math.round(x0 + m), top: Math.round(y0 + m), width: Math.max(4, Math.round(x1 - x0 - 2 * m)), height: Math.max(4, Math.round(y1 - y0 - 2 * m)) });
  const ROC_RE = /(1\d{2})\D{1,2}(\d{1,2})\D{1,2}(\d{1,2})/;
  const han = (t) => String(t || '').normalize('NFKC').replace(/[^\p{Script=Han}-]/gu, '');

  /** 辨識前清除水平線（底線、格線）：長度超過 minLen 的連續暗點列塗白 */
  function eraseHLines(ctx, x0, y0, x1, y1, minLen) {
    const w = x1 - x0; const h = y1 - y0;
    const img = ctx.getImageData(x0, y0, w, h);
    const d = img.data;
    const dark = (i) => 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2] < 170;
    for (let y = 0; y < h; y += 1) {
      let run = 0;
      for (let x = 0; x <= w; x += 1) {
        if (x < w && dark((y * w + x) * 4)) { run += 1; continue; }
        if (run >= minLen) for (let k = x - run; k < x; k += 1) { const i = (y * w + k) * 4; d[i] = 255; d[i + 1] = 255; d[i + 2] = 255; }
        run = 0;
      }
    }
    ctx.putImageData(img, x0, y0);
  }

  /** 多行文字格：依水平投影切成一行一行，各自以單行模式辨識（Word 版學員欄常有 2～3 行） */
  async function ocrLines(w, canvas, ctx, r) {
    const d = ctx.getImageData(r.left, r.top, r.width, r.height).data;
    const rowInk = [];
    for (let y = 0; y < r.height; y += 1) {
      let n = 0;
      for (let x = 0; x < r.width; x += 1) { const p = (y * r.width + x) * 4; if (0.299 * d[p] + 0.587 * d[p + 1] + 0.114 * d[p + 2] < 150) n += 1; }
      rowInk.push(n > 2);
    }
    const segs = []; let st = -1;
    rowInk.forEach((v, y) => { if (v && st < 0) st = y; if (!v && st >= 0) { if (y - st > 12) segs.push([st, y]); st = -1; } });
    if (st >= 0 && r.height - st > 12) segs.push([st, r.height]);
    const out = [];
    for (const [a, b] of segs) {
      const rr = { left: r.left, top: r.top + Math.max(0, a - 6), width: r.width, height: Math.min(r.height - a, b - a + 12) };
      if (window.PaddleRec) {
        // PaddleOCR（中文印刷／手寫都較準）：切出這一行影像再辨識
        const c = document.createElement('canvas'); c.width = rr.width; c.height = rr.height;
        c.getContext('2d').drawImage(canvas, rr.left, rr.top, rr.width, rr.height, 0, 0, rr.width, rr.height);
        try { out.push((await PaddleRec.recognize(c)).text); continue; } catch (e) { /* 失敗時改用 Tesseract */ }
      }
      out.push((await w.recognize(canvas, { rectangle: rr })).data.text);
    }
    return out.join(' ');
  }

  /** 淡紫色螢光筆塗色（內科部門診表上標示要申請的門診） */
  const isHighlight = (R, G, B) => B - G > 25 && B - R > 5 && B > 160 && R > 100 && G > 110;

  /** 找出塗色區塊（以 6px 網格合併相鄰塗色點），回傳區塊外框 */
  function highlightBoxes(ctx, w, h) {
    const d = ctx.getImageData(0, 0, w, h).data;
    const step = 6; const gw = Math.ceil(w / step); const gh = Math.ceil(h / step);
    const grid = new Uint8Array(gw * gh);
    for (let gy = 0; gy < gh; gy += 1) for (let gx = 0; gx < gw; gx += 1) {
      const p = ((gy * step) * w + gx * step) * 4;
      if (isHighlight(d[p], d[p + 1], d[p + 2])) grid[gy * gw + gx] = 1;
    }
    const seen = new Uint8Array(gw * gh); const boxes = [];
    for (let i = 0; i < grid.length; i += 1) {
      if (!grid[i] || seen[i]) continue;
      let x0 = 1e9; let y0 = 1e9; let x1 = 0; let y1 = 0; let n = 0; const st = [i]; seen[i] = 1;
      while (st.length) {
        const k = st.pop(); const x = k % gw; const y = (k - x) / gw; n += 1;
        x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y);
        for (let dy = -2; dy <= 2; dy += 1) for (let dx = -2; dx <= 2; dx += 1) {
          const nx = x + dx; const ny = y + dy; const j = ny * gw + nx;
          if (nx >= 0 && ny >= 0 && nx < gw && ny < gh && grid[j] && !seen[j]) { seen[j] = 1; st.push(j); }
        }
      }
      if (n >= 8) boxes.push({ x0: x0 * step, y0: y0 * step, x1: (x1 + 1) * step, y1: (y1 + 1) * step });
    }
    return boxes;
  }

  const CODES = ['GM', 'CV', 'CM', 'HEMA', 'ONC', 'INF', 'META', 'NEPH', 'NEUR', 'AIR', 'GI', 'GER'];
  const VALID_CODE = new RegExp(`^(${CODES.join('|')})$`);
  const levD = (a, b) => {
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
    for (let j = 1; j <= b.length; j += 1) d[0][j] = j;
    for (let i = 1; i <= a.length; i += 1) for (let j = 1; j <= b.length; j += 1) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[a.length][b.length];
  };
  /** OCR 代碼 → 最接近的有效代碼（差 1 個字母以內、且唯一） */
  function nearestCode(raw) {
    if (!raw) return '';
    if (VALID_CODE.test(raw)) return raw;
    const sc = CODES.map((c) => ({ c, d: levD(raw, c) })).sort((a, b) => a.d - b.d);
    return sc[0].d <= 1 && (!sc[1] || sc[1].d > sc[0].d) ? sc[0].c : raw;
  }
  /** 讀排班表的塗色格：代碼（格內）＋指導醫師（下一格） */
  async function parseSchedule(P, canvas) {
    const ctx = canvas.getContext('2d');
    const items = [];
    for (const b of highlightBoxes(ctx, canvas.width, canvas.height)) {
      const bh = b.y1 - b.y0;
      // 代碼：塗色底的英文字 Tesseract 常讀不出，先用 PaddleOCR，失敗再用 Tesseract
      const cr = rect(b.x0 - 6, b.y0 - 6, b.x1 + 6, b.y1 + 6);
      let raw = '';
      if (window.PaddleRec) {
        const c = document.createElement('canvas'); c.width = cr.width; c.height = cr.height;
        c.getContext('2d').drawImage(canvas, cr.left, cr.top, cr.width, cr.height, 0, 0, cr.width, cr.height);
        try { raw = (await PaddleRec.recognize(c)).text.normalize('NFKC').toUpperCase().replace(/[^A-Z]/g, ''); } catch (e) { raw = ''; }
      }
      if (!VALID_CODE.test(nearestCode(raw))) {
        const t = ((await P.lat.recognize(canvas, { rectangle: cr })).data.text || '').replace(/[^A-Z]/g, '');
        if (VALID_CODE.test(nearestCode(t)) || !raw) raw = t;
      }
      const code = nearestCode(raw);
      // 指導醫師在緊接著的下一格（約一個格子高）：用 PaddleOCR 讀，之後以姓名相似度比對 rawdata 的老師
      const below = rect(b.x0 - 15, b.y1 + 2, b.x1 + 15, b.y1 + 2 + Math.max(bh * 1.3, 50));
      let teacher = '';
      if (window.PaddleRec) {
        const c = document.createElement('canvas'); c.width = below.width; c.height = below.height;
        c.getContext('2d').drawImage(canvas, below.left, below.top, below.width, below.height, 0, 0, below.width, below.height);
        try { teacher = han((await PaddleRec.recognize(c)).text); } catch (e) { teacher = ''; }
      }
      if (teacher.length < 2) teacher = han((await P.chi7.recognize(canvas, { rectangle: below })).data.text);
      items.push({ code, raw, teacher, box: [b.x0, b.y0, b.x1, b.y1] });
    }
    // 依門診表的時間順序排列：先分週（同一橫列帶），再依星期／時段（由左到右）
    const bandH = Math.max(40, ...items.map((x) => x.box[3] - x.box[1])) * 2.5;
    items.sort((a, b) => (Math.abs(a.box[1] - b.box[1]) < bandH ? a.box[0] - b.box[0] : a.box[1] - b.box[1]));
    items.forEach((x, i) => { x.seq = i; });
    // 有效代碼的「種類數」（印章等雜訊常被重複讀成同一個代碼）
    return { items, valid: new Set(items.filter((x) => VALID_CODE.test(x.code)).map((x) => x.code)).size };
  }

  function inkRatio(ctx, r) {
    const d = ctx.getImageData(r.left, r.top, r.width, r.height).data;
    let n = 0;
    for (let p = 0; p < d.length; p += 4) if (0.299 * d[p] + 0.587 * d[p + 1] + 0.114 * d[p + 2] < 150) n += 1;
    return n / (d.length / 4);
  }

  async function renderPage(page, extraRotation, dpi) {
    const viewport = page.getViewport({ scale: dpi / 72, rotation: (page.rotate + extraRotation) % 360 });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width; canvas.height = viewport.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    return canvas;
  }
  const findGrid = (ctx, w, h) => { const img = ctx.getImageData(0, 0, w, h); return SheetDetect.detectGrid(img) || SheetDetect.detectGrid(img, 0.15); };

  /** 自動轉正：試 0°／90°／270°，取表格前幾列中文字＋數字最多的方向；都沒有表格回傳 null */
  async function chooseRotation(P, page) {
    let best = null; let withGrid = 0;
    for (const rot of [0, 90, 270]) {
      const cv = await renderPage(page, rot, 150);
      const grid = findGrid(cv.getContext('2d'), cv.width, cv.height);
      if (grid && grid.rows.length >= 2) {
        withGrid += 1;
        const r0 = grid.rows[Math.min(1, grid.rows.length - 1)]; const r1 = grid.rows[Math.min(grid.rows.length - 1, 4)];
        const t = (await P.chi6.recognize(cv, { rectangle: rect(grid.cols[0], r0.top, grid.cols[grid.cols.length - 1], r1.bottom) })).data.text.normalize('NFKC');
        const score = (t.match(/\p{Script=Han}/gu) || []).length + (t.match(/\d/g) || []).length * 0.5;
        if (!best || score > best.score) best = { rot, score };
      }
      cv.width = 0;
      if (rot === 0 && best && best.score >= 20) break;   // 正向就很清楚，不必再試
    }
    return best && (best.score >= 6 || withGrid === 1) ? best.rot : null;
  }

  function create({ root, getRawRows, getYm, onChange }) {
    let pages = [];   // { key, file, page, unit, note, rot, rows: [...] }
    let seq = 0;
    let busy = false;
    const queue = [];
    let matched = [];

    root.innerHTML = '<div class="ap-status muted small"></div><div class="stats ap-stats"></div><div class="ap-notes small muted"></div>';
    const $ = (s) => root.querySelector(s);
    const setStatus = (m) => { $('.ap-status').textContent = m; };

    async function addPdf(file) {
      const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false }).promise;
      for (let p = 1; p <= pdf.numPages; p += 1) {
        const pg = { key: ++seq, file: file.name, page: p, unit: '', note: '', rows: [] };
        pages.push(pg);
        queue.push({ pg, page: await pdf.getPage(p) });
      }
      render();
      run();
    }

    async function run() {
      if (busy) return;
      busy = true;
      try {
        const workers = await getPool(setStatus);
        let done = pages.filter((p) => p.done).length;
        await Promise.all(workers.map(async (P) => {
          while (queue.length) {
            const job = queue.shift();
            setStatus(`辨識申請單：${done + 1}/${pages.length} 頁…`);
            // 低解析度判斷不出方向時（細格線在 150 dpi 可能偵測不到），以 300 dpi 正向再試
            const rot = await chooseRotation(P, job.page);
            job.pg.rot = rot ?? 0;
            const canvas = await renderPage(job.page, job.pg.rot, 300);
            await extractPage(P, job.pg, canvas, job.page);
            canvas.width = 0;
            if (rot === null && !job.pg.rows.length && !job.pg.unreadable) job.pg.note = '沒有偵測到申請資料（可能是排班表或空白頁），略過';
            job.pg.done = true;
            done += 1;
            rematch();
          }
        }));
        busy = false;
        setStatus(`辨識完成：${pages.length} 頁、${pages.reduce((a, p) => a + p.rows.length, 0)} 列。`);
        rematch();
      } catch (err) {
        console.error(err);
        setStatus(`文字辨識失敗：${err.message}`);
      } finally {
        busy = false;
      }
    }

    /** 住診教學紀錄表（手寫）：一頁一筆，讀「指導教師」與「教學起迄時間」的日期 */
    async function extractRecordForm(P, pg, canvas, text) {
      const t = text.normalize('NFKC').replace(/\s/g, '');
      pg.format = '紀錄表';
      const teacher = ((t.match(/指導教師[:：]?([\p{Script=Han}]{2,4}?)醫師/u) || t.match(/指導教師[:：]?([\p{Script=Han}]{2,4})/u) || [])[1] || '');
      const unit = ((t.match(/科別[:：]?([\p{Script=Han}]{2,8}?)指導/u) || [])[1] || '');
      const d = t.match(/(\d{2,3})年(\d{1,2})月(\d{1,2})日/);
      const dateText = d ? `${Number(d[1]) + 1911}-${d[2]}-${d[3]}` : '';
      if (unit) pg.unit = unit;
      if (teacher.length >= 2 || dateText) {
        pg.rows.push({ id: ++seq, ocrText: `${dateText} ${teacher}`, ocrDate: dateText,
          date: A.fixDate(dateText, getYm() || '2000-01'), teacher, student: '', include: true, card: '', order: pg.key * 1000 });
      } else pg.unreadable = (pg.unreadable || 0) + 1;
    }

    async function extractPage(P, pg, canvas, page) {
      const ctx = canvas.getContext('2d');
      // 內科部門診表：有塗色格且讀得到科別代碼 → 讀「科別代碼」與下方「指導醫師」，不當申請單列處理
      const nBoxes = highlightBoxes(ctx, canvas.width, canvas.height).length;
      if (nBoxes >= 4) {
        let best = await parseSchedule(P, canvas);
        // 門診表常是橫向掃描：塗色區塊夠多（≥6，印章通常只有 2～4 塊）才試其他方向，取讀到最多有效代碼者
        if (best.valid < 4 && page && nBoxes >= 6) {
          for (const extra of [90, 270, 180]) {
            const cv = await renderPage(page, ((pg.rot || 0) + extra) % 360, 300);
            const res = await parseSchedule(P, cv);
            cv.width = 0;
            if (res.valid > best.valid) best = res;
          }
        }
        if (best.valid >= 4) {
          pg.format = '內科部門診表';
          pg.schedule = best.items.filter((x) => VALID_CODE.test(x.code));
          pg.scheduleAll = best.items;   // 含無效代碼（除錯用）
          pg.note = `內科部門診表：塗色 ${pg.schedule.length} 格（${pg.schedule.map((x) => `${x.code} ${x.teacher}`).join('、')}）`;
          return;
        }
      }
      // 標題判斷：「住診教學紀錄表」（手寫，一頁一筆）
      const top = (await P.chi6.recognize(canvas, { rectangle: rect(0, 0, canvas.width, canvas.height * 0.3) })).data.text;
      if (/紀錄表|記錄表/.test(top.replace(/\s/g, ''))) { await extractRecordForm(P, pg, canvas, top); return; }
      const grid = findGrid(ctx, canvas.width, canvas.height);
      if (!grid || grid.cols.length < 4) { pg.note = '沒有偵測到表格'; return; }
      const c = grid.cols;
      const head = grid.rows[0];
      const data = grid.rows.slice(1);
      // 單位：表格上方文字中的「單位:XXX」（排除「單位簽章」）
      const headTop = Math.max(0, head.top - canvas.height * 0.2);
      const headText = (await P.chi6.recognize(canvas, { rectangle: rect(0, headTop, canvas.width, head.top) })).data.text;
      const units = [...headText.normalize('NFKC').replace(/\s/g, '').matchAll(/單位(?!簽章)[:：]?([\p{Script=Han}]{2,10}?)(?=單位|申請|主治|$)/gu)].map((m) => m[1]);
      pg.unit = units[units.length - 1] || '';

      // 清除水平線（文字常壓在格線／底線上）
      const xEnd = c[c.length - 1];
      eraseHLines(ctx, c[0], Math.max(0, head.top - 6), xEnd, Math.min(canvas.height, data[data.length - 1].bottom + 6), Math.round((xEnd - c[0]) * 0.06));
      const cell = (k, r, w) => w.recognize(canvas, { rectangle: rect(c[k] + 8, r.top + 4, c[k + 1] - 8, Math.min(canvas.height, r.bottom + 4)) }).then((x) => x.data.text);
      // 版面：表頭第一格是「主治醫師…」→ Word 版；否則系統印出版
      const h0 = han(await cell(0, head, P.chi7));
      const word = /主治|醫師/.test(h0);
      pg.format = word ? 'Word 版' : '系統版';

      for (const r of data) {
        if (inkRatio(ctx, rect(c[0], r.top, c[Math.min(3, c.length - 1)], r.bottom, 8)) < 0.004) continue;   // 空白列
        let dateText; let teacher; let students;
        if (word) {
          const [t, s, d] = await Promise.all([cell(0, r, P.chi7), ocrLines(P.chi6, canvas, ctx, rect(c[1] + 8, r.top + 4, c[2] - 8, r.bottom)), cell(2, r, P.num)]);
          teacher = han(t);
          students = s.normalize('NFKC').split(/[、,，;；\s\n]+/).map(han).filter((x) => x.length >= 2);
          const m = d.normalize('NFKC').match(ROC_RE);
          dateText = m ? `${Number(m[1]) + 1911}-${m[2]}-${m[3]}` : d;
        } else {
          const [d, t, s] = await Promise.all([cell(0, r, P.num), cell(1, r, P.chi7), cell(2, r, P.chi7)]);
          dateText = d; teacher = han(t); students = [han(s)];
        }
        teacher = teacher.replace(/醫師$/, '');
        // 過濾雜訊列（排班表、手寫塗改、空白表單邊角）：需有 2 字以上的主治醫師，且有日期或學員
        const hasDate = !!A.fixDate(dateText, getYm() || '2000-01');
        if (teacher.length < 2 || !(hasDate || students.some((x) => x.length >= 2))) {
          // 有墨跡但讀不出：可能是手寫或塗改 → 記錄，輸出紅字請人工核對
          if (inkRatio(ctx, rect(c[0], r.top, c[Math.min(3, c.length - 1)], r.bottom, 8)) >= 0.02) pg.unreadable = (pg.unreadable || 0) + 1;
          continue;
        }
        // 系統印出版：申請單上「老師完成否」欄（第 6 欄）有沒有 Y；主治醫師格右側空白處有沒有審核畫記「/」
        let teacherDone = null; let marked = false;
        if (!word && c.length >= 8) teacherDone = inkRatio(ctx, rect(c[5], r.top, c[6], r.bottom, 10)) > 0.008;
        if (!word) {
          const tw = c[2] - c[1];
          marked = inkRatio(ctx, rect(c[2] - tw * 0.3, r.top, c[2] - 6, r.bottom, 4)) > 0.01;
        }
        (students.length ? students : ['']).forEach((st) => pg.rows.push({
          id: ++seq, ocrText: `${dateText} ${teacher} ${st}`.trim(), ocrDate: dateText,
          date: A.fixDate(dateText, getYm() || '2000-01'), teacher, student: st, include: true, card: '',
          word, teacherDone, marked, order: pg.key * 1000 + pg.rows.length,   // 彙整表出現順序（跨檔案依匯入順序）
        }));
      }
      pg.hasMarks = pg.rows.some((x) => x.marked);
    }

    function rematch() {
      const ym = getYm();
      const raw = getRawRows();
      const flat = pages.flatMap((p) => p.rows.map((r) => ({ ...r, unit: p.unit, page: p })));
      flat.forEach((r) => { if (ym) r.date = A.fixDate(r.ocrDate, ym) || r.date; });
      matched = ym && raw.length ? A.matchRows(flat, raw, ym) : flat.map((r) => ({ ...r, match: null, status: 'no-raw' }));
      render();
      onChange();
    }

    function render() {
      const count = (s) => matched.filter((m) => m.status === s).length;
      const undated = matched.filter((m) => !m.match && !m.date).length;
      $('.ap-stats').innerHTML = matched.length ? `
        <div class="stat"><span>申請單列數</span><b>${matched.length}</b></div>
        <div class="stat"><span>與 rawdata 對到</span><b>${count('ok') + count('unsure') + count('not-done')}</b></div>
        <div class="stat"><span>其中老師未評核</span><b>${count('not-done')}</b></div>
        <div class="stat"><span>系統無紀錄（紅字計入）</span><b class="text-red">${count('no-system') - undated}</b></div>
        <div class="stat"><span>日期無法辨識（紅字、金額空白）</span><b class="text-red">${undated}</b></div>
        <div class="stat"><span>整列無法辨識（手寫等，紅字提醒）</span><b class="text-red">${pages.reduce((a, p) => a + (p.unreadable || 0), 0)}</b></div>` : '';
      const sched = pages.filter((p) => p.schedule);
      const notes = pages.filter((p) => p.note && !p.schedule);
      $('.ap-notes').innerHTML = [
        sched.map((p) => `${esc(p.file)} 第 ${p.page} 頁 ${esc(p.note)}`).join('<br>'),
        notes.length ? `略過的頁：${notes.map((p) => `${esc(p.file)} 第 ${p.page} 頁`).join('、')}` : '',
      ].filter(Boolean).join('<br>');
    }

    const api = {
      addPdf,
      rematch,
      hasData: () => pages.length > 0,
      busy: () => busy || queue.length > 0,
      matched: () => matched,
      /** 有日期的申請單列 → 計費紀錄（系統無紀錄者 inSystem=false、紅字） */
      records: () => A.toRecords(matched.filter((m) => m.match || m.date), getRawRows()),
      /** 日期無法辨識的列（輸出為紅字、金額空白，供人工核對） */
      undated: () => [
        ...matched.filter((m) => !m.match && !m.date).map((m) => ({ unit: m.unit, teacher: m.teacher, student: m.student, where: `${m.page.file} p${m.page.page}`, text: m.ocrText })),
        // 整列讀不出（手寫等）：每頁一列紅字提醒
        ...pages.filter((p) => p.unreadable).map((p) => ({ unit: p.unit, teacher: `（${p.unreadable} 列無法辨識：${p.file} 第${p.page}頁，請人工核對）`, student: '', where: `${p.file} p${p.page}`, text: '' })),
      ],
      /** 內科部門診表（塗色格）：[{ file, page, code, teacher }] */
      schedule: () => pages.flatMap((p) => (p.schedule || []).map((x) => ({ file: p.file, page: p.page, ...x }))),
      /** 與申請單對到的 rawdata 列號（綠色標示用） */
      matchedRowNums: () => matched.filter((m) => m.match && m.match.rowNum).map((m) => m.match.rowNum),
      reset() { pages = []; queue.length = 0; matched = []; render(); setStatus(''); },
    };
    instances.push(api);   // 測試／除錯用
    return api;
  }
  const instances = [];

  /**
   * 下載比對後的 rawdata：與申請單對到的列，「評核者卡號」「評核者」兩格標綠色（只改顏色，其餘原樣）
   * raw: { name, buffer }；rowNums: 原始 Excel 列號
   */
  async function downloadMatchedRaw(raw, rowNums) {
    const wb = XLSX.read(raw.buffer, { sheets: ['課程rawdata'], sheetRows: 1 });
    const header = XLSX.utils.sheet_to_json(wb.Sheets['課程rawdata'], { header: 1 })[0] || [];
    const cols = ['評核者卡號', '評核者'].map((h) => header.indexOf(h) + 1).filter((c) => c > 0);
    if (cols.length < 2) throw new Error('rawdata 找不到「評核者卡號」「評核者」欄位');
    const writes = [...new Set(rowNums)].flatMap((row) => cols.map((col) => ({ sheet: '課程rawdata', row, col, kind: 'style', fill: 'FF92D050' })));
    const out = await TutorFill.patchXlsx(raw.buffer, writes, { JSZip: window.JSZip, DOMParser, XMLSerializer });
    const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = `${raw.name.replace(/\.xlsx$/i, '')}_比對後.xlsx`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    return writes.length / cols.length;
  }

  window.AppPdf = { create, downloadMatchedRaw, instances };
})();
