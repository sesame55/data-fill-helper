/*
 * 分頁：全人電影（心得張貼 ＋ 票選資料 → 心得票數表）
 * 計算邏輯在 js/core/film-calc.js。所有資料只存在記憶體（state），清空、關閉分頁即消失。
 */
(() => {
  const F = window.FilmCalc;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const emptyState = () => ({ essays: null, votes: null, result: null, busy: [] });
  let state = emptyState();

  const dropzone = $('fm-dropzone');
  const input = $('fm-file');
  input.addEventListener('change', () => { handleFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
  dropzone.addEventListener('drop', (e) => handleFiles([...e.dataTransfer.files]));

  async function handleFiles(files) {
    for (const f of files) {
      setBusy(f.name, true);
      try {
        if (!/\.xlsx$/i.test(f.name)) throw new Error('只接受 .xlsx');
        const wb = XLSX.read(await f.arrayBuffer());
        const sheets = wb.SheetNames.map((name) => ({ name, rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' }) }));
        const type = F.detectType(sheets);
        if (type === 'essays') {
          const data = F.parseEssays(sheets);
          if (!data.essays.length) throw new Error('找不到任何心得（A 編號、B 心得、C 作者）');
          state.essays = { name: f.name, data };
        } else if (type === 'votes') {
          state.votes = { name: f.name, data: F.parseVotes(sheets) };
        } else {
          throw new Error(`看不出是心得張貼檔（工作表名稱如 101-200）或票選資料檔（含「${F.VOTE_SHEET}」工作表）`);
        }
      } catch (err) {
        console.error(err);
        alert(`讀取「${f.name}」失敗：${err.message}`);
      } finally {
        setBusy(f.name, false);
      }
    }
    recompute();
  }

  function setBusy(name, on) {
    state.busy = on ? [...state.busy, name] : state.busy.filter((n) => n !== name);
    renderStatus();
  }

  const keywords = () => $('fm-keywords').value.split(/[、,，\s]+/).filter(Boolean);

  function recompute() {
    state.result = state.essays && state.votes ? F.analyze(state.essays.data, state.votes.data, keywords()) : null;
    renderStatus();
    renderResult();
  }

  function renderStatus() {
    const item = (ok, label, detail) =>
      `<div class="status-item ${ok ? 'ok' : ''}"><span class="dot"></span><b>${label}</b><span class="muted small">${detail}</span></div>`;
    const e = state.essays;
    const v = state.votes;
    $('fm-status').innerHTML = [
      item(!!e, '心得張貼', e ? `${esc(e.name)}・${e.data.essays.length} 篇（${e.data.essays[0].no}～${e.data.essays.at(-1).no}）` : '尚未匯入'),
      item(!!v, '票選資料', v ? `${esc(v.name)}・${v.data.ballots.length.toLocaleString()} 張票` : '尚未匯入'),
      state.busy.length ? `<div class="status-item busy"><span class="dot"></span>讀取中：${esc(state.busy.join('、'))}</div>` : '',
    ].join('');
  }

  const kwLabel = (r) => r.keywords.join('/') || '關鍵字';

  function renderResult() {
    const r = state.result;
    $('fm-result-step').hidden = !r;
    if (!r) return;
    const s = r.stats;
    const m = r.match;
    $('fm-stats').innerHTML = `
      <div class="stat"><span>心得</span><b>${r.rows.length}</b></div>
      <div class="stat"><span>投票張數</span><b>${s.ballots.toLocaleString()}</b></div>
      <div class="stat"><span>有效票</span><b>${s.validVotes.toLocaleString()}</b></div>
      <div class="stat muted-stat"><span>全部無效的票</span><b>${s.noValid}</b></div>
      <div class="stat"><span>部門含「${esc(kwLabel(r))}」</span><b>${m.flagged} 位</b></div>
      <div class="stat"><span>作者 02 查無此人</span><b class="${m.notFound ? 'text-red' : ''}">${m.notFound}</b></div>`;
    const warns = [];
    if (state.essays.data.skipped.length) warns.push(`略過沒有作者欄的工作表：${state.essays.data.skipped.join('、')}`);
    if (s.dupBallots) warns.push(`${s.dupBallots} 張票有重複號碼，只算一次`);
    if (s.extraBallots) warns.push(`${s.extraBallots} 張票填超過 3 個有效號碼，只算前 3 個`);
    if (m.notFound) warns.push(`${m.notFound} 位作者在票選資料找不到（沒投票或姓名寫法不同），無法判斷部門`);
    if (m.multi) warns.push(`${m.multi} 位作者同名多筆、部門不只一個，請人工確認`);
    if (m.backup) warns.push(`${m.backup} 位作者不在「${F.VOTE_SHEET}」，改由其他名單工作表對照部門`);
    if (m.noAuthor) warns.push(`${m.noAuthor} 篇心得沒有作者姓名`);
    $('fm-warnings').innerHTML = warns.map((w) => `<div class="warn">⚠ ${esc(w)}</div>`).join('');
    $('fm-flag-head').textContent = kwLabel(r);
    $('fm-only-flag-label').textContent = `只顯示部門含「${kwLabel(r)}」`;
    renderTable();
    renderBallots();
  }

  function renderTable() {
    const r = state.result;
    if (!r) return;
    const q = $('fm-filter').value.trim();
    const onlyFlag = $('fm-only-flag').checked;
    const rows = r.rows.filter((x) => (!onlyFlag || x.flag) && (!q || [x.no, x.author, x.dept, x.text, x.note].some((v) => String(v).includes(q))));
    $('fm-result').innerHTML = rows.map((x) => `<tr class="${x.flag ? 'row-warn' : ''}">
      <td>${x.no}</td><td class="film-text" title="${esc(x.text)}"><div class="clamp">${esc(x.text)}</div></td><td>${esc(x.author)}</td>
      <td class="num">${x.votes}</td><td>${esc(x.dept)}</td><td>${x.flag ? '是' : ''}</td>
      <td class="${x.note ? 'text-red' : ''} small">${esc(x.note)}</td></tr>`).join('')
      || '<tr class="empty"><td colspan="7">沒有符合的資料</td></tr>';
  }

  function renderBallots() {
    const list = state.result.detail.filter((d) => d.note);
    $('fm-ballot-count').textContent = list.length;
    $('fm-ballots').innerHTML = list.map((d) => `<tr><td>${d.row}</td><td>${esc(d.acct)}</td><td>${esc(d.name)}</td>
      <td class="film-raw">${esc(d.raw)}</td><td>${d.valid.join('、')}</td><td class="text-red small">${esc(d.note)}</td></tr>`).join('');
  }

  $('fm-filter').addEventListener('input', renderTable);
  $('fm-only-flag').addEventListener('change', renderTable);
  $('fm-keywords').addEventListener('change', recompute);

  /* ---------- 下載 ---------- */
  function styleSheet(ws, widths) {
    const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
    ws.columns = widths.map((width) => ({ width }));
    ws.eachRow((row, i) => row.eachCell({ includeEmpty: true }, (c) => {
      c.border = { top: thin, bottom: thin, left: thin, right: thin };
      c.alignment = { vertical: i === 1 ? 'middle' : 'top', horizontal: i === 1 ? 'center' : undefined, wrapText: true };
      if (i === 1) { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }; }
    }));
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: widths.length } };
  }

  $('fm-download').addEventListener('click', async () => {
    const r = state.result;
    if (!r) return;
    const btn = $('fm-download');
    btn.disabled = true;
    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('心得票數');
      ws.addRow(['編號', '心得', '心得作者', '票數', '部門（對照02）', kwLabel(r), '對照說明']);
      const hl = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
      for (const x of r.rows) {
        const row = ws.addRow([x.no, x.text, x.author, x.votes, x.dept || null, x.flag ? '是' : null, x.note || null]);
        if (x.flag) for (let c = 1; c <= 7; c++) row.getCell(c).fill = hl;
        if (x.note) row.getCell(7).font = { color: { argb: 'FFC00000' } };
      }
      styleSheet(ws, [8, 90, 12, 8, 22, 10, 20]);
      ws.getColumn(1).numFmt = '@';
      [1, 3, 4, 6].forEach((c) => ws.getColumn(c).eachCell((cell, i) => { if (i > 1) cell.alignment = { ...cell.alignment, horizontal: 'center' }; }));

      const d = wb.addWorksheet('計票明細');
      d.addRow(['原始列號', '帳號', '姓名', '原始填答', '有效票1', '有效票2', '有效票3', '處理說明']);
      for (const b of r.detail) {
        const row = d.addRow([b.row, b.acct, b.name, b.raw, b.valid[0] ?? null, b.valid[1] ?? null, b.valid[2] ?? null, b.note || null]);
        if (b.note) row.getCell(8).font = { color: { argb: 'FFC00000' } };
      }
      styleSheet(d, [9, 9, 10, 40, 8, 8, 8, 40]);

      const out = await wb.xlsx.writeBuffer();
      const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url;
      const film = (state.essays.name.match(/[-－]\s*([^-－]+)\.xlsx$/i) || [])[1];
      a.download = `心得票數${film ? `-${film.trim()}` : ''}.xlsx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error(err);
      alert(`產生檔案失敗：${err.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  App.registerTab('film', {
    reset() { state = emptyState(); $('fm-filter').value = ''; recompute(); },
  });
  renderStatus();
})();
