/*
 * 分頁：導生座談 → 導談完成率（rawdata 填入原始統計檔）
 * 只修改當月空白格（直接改 xlsx 內的 XML），原檔其他內容、格式、條件式格式皆保留。
 * 所有資料只存在記憶體（state），清空、關閉分頁即消失。
 */
(() => {
  const T = window.TutorCalc;
  const F = window.TutorFill;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const mmdd = (iso) => (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso) ? `${iso.slice(5, 7)}/${iso.slice(8, 10)}` : (iso || ''));

  const emptyState = () => ({ raws: {}, book: null, report: null, busy: [] });
  let state = emptyState();

  const dropzone = $('tf-dropzone');
  const input = $('tf-file');
  input.addEventListener('change', () => { handleFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach((ev) => dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
  dropzone.addEventListener('drop', (e) => handleFiles([...e.dataTransfer.files]));

  async function handleFiles(files) {
    for (const f of files) {
      setBusy(f.name, true);
      try {
        if (!/\.xlsx$/i.test(f.name)) throw new Error('只接受 .xlsx');
        const buffer = await f.arrayBuffer();
        const wb = XLSX.read(buffer, { cellDates: true, sheets: 0 });
        const first = wb.Sheets[wb.SheetNames[0]];
        const type = T.detectRawType(XLSX.utils.sheet_to_json(first, { header: 1 })[0] || []);
        if (type) {
          state.raws[type] = { name: f.name, records: T.parseRaw(XLSX.utils.sheet_to_json(first, { defval: '', raw: false }), type) };
        } else {
          const ewb = new ExcelJS.Workbook();
          await ewb.xlsx.load(buffer.slice(0));
          if (!ewb.worksheets.some((ws) => /^(R檢核|PGY\(|C\()/.test(ws.name))) throw new Error('看不出是 rawdata 或導談完成率統計');
          state.book = { name: f.name, buffer, wb: ewb };
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

  function recompute() {
    const ym = $('tf-month').value;
    const rawAll = Object.values(state.raws).flatMap((r) => r.records);
    state.report = state.book && rawAll.length ? F.planFill(state.book.wb, rawAll, ym ? [ym] : undefined) : null;
    renderStatus();
    renderReport();
  }

  function renderStatus() {
    const item = (ok, label, detail) =>
      `<div class="status-item ${ok ? 'ok' : ''}"><span class="dot"></span><b>${label}</b><span class="muted small">${detail}</span></div>`;
    const raw = (cat, label) => item(!!state.raws[cat], `rawdata ${label}`,
      state.raws[cat] ? `${esc(state.raws[cat].name)}・${state.raws[cat].records.length.toLocaleString()} 筆` : '尚未匯入');
    $('tf-status').innerHTML = [
      raw('PGY', 'PGY'), raw('R', '住院醫師'), raw('UGY', '實習生'),
      item(!!state.book, '導談完成率統計', state.book ? esc(state.book.name) : '尚未匯入'),
      state.busy.length ? `<div class="status-item busy"><span class="dot"></span>讀取中：${esc(state.busy.join('、'))}</div>` : '',
    ].join('');
  }

  function renderReport() {
    const r = state.report;
    $('tf-result-step').hidden = !r;
    if (!r) return;
    const diff = r.kept.filter((k) => !k.same);
    $('tf-stats').innerHTML = `
      <div class="stat"><span>本次填入</span><b>${r.filled.length} 人</b></div>
      <div class="stat muted-stat"><span>已有資料且一致</span><b>${r.kept.length - diff.length}</b></div>
      <div class="stat"><span>已有資料但不同（未覆寫）</span><b class="${diff.length ? 'text-red' : ''}">${diff.length}</b></div>
      <div class="stat"><span>名冊找不到</span><b class="${r.notInRoster.length ? 'text-red' : ''}">${r.notInRoster.length}</b></div>`;
    const warns = [...r.missing];
    ['PGY', 'R', 'UGY'].filter((c) => !state.raws[c]).forEach((c) => warns.push(`尚未匯入 ${T.CAT_LABEL[c]} rawdata`));
    $('tf-warnings').innerHTML = warns.map((w) => `<div class="warn">⚠ ${esc(w)}</div>`).join('');
    const roc = (ym) => `${Number(ym.slice(0, 4)) - 1911}/${ym.slice(5, 7)}`;
    $('tf-sheets').innerHTML = r.sheets.map((s) => `<tr><td>${roc(s.ym)}</td><td>${esc(s.name)}</td>
      <td>${s.statusCol ? `${s.statusCol} 欄起` : '<span class="problem">找不到</span>'}</td>
      <td class="num">${s.filled}</td><td class="num">${s.kept}</td><td class="num ${s.skipped ? 'text-red' : ''}">${s.skipped}</td></tr>`).join('');

    const diffRow = (x) => `<tr><td class="small">${esc(x.sheet)}</td><td>${x.row}</td><td>${esc(x.card)}</td><td>${esc(x.name)}</td><td>${esc(x.existing)} ${mmdd(x.existingDate)}</td><td>V ${mmdd(x.date)}</td></tr>`;
    const table = (title, head, rows) => (rows.length ? `<h4>${title}（${rows.length}）</h4><div class="table-wrap"><table class="file-table">
      <thead><tr>${head.map((h) => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.join('')}</tbody></table></div>` : '');
    $('tf-lists').innerHTML = [
      table('本次填入', ['工作表', '列', '卡號', '姓名', '填入'], r.filled.map((x) => `<tr><td class="small">${esc(x.sheet)}</td><td>${x.row}</td><td>${esc(x.card)}</td><td>${esc(x.name)}</td><td>V ${mmdd(x.date)}${x.replaced ? '<span class="muted small">（覆蓋「算完成」）</span>' : ''}${x.count > 1 ? `<span class="muted small">（本月 ${x.count} 筆，取最早）</span>` : ''}</td></tr>`)),
      table('表上已是 V 但日期不同（未覆寫，請確認）', ['工作表', '列', '卡號', '姓名', '表上現有', '系統紀錄'], diff.filter((x) => x.existing === 'V' || !x.existing).map(diffRow)),
      table('表上已手動填寫（院內／院外／外訓…，未覆寫，供參考）', ['工作表', '列', '卡號', '姓名', '表上現有', '系統紀錄'], diff.filter((x) => x.existing && x.existing !== 'V').map(diffRow)),
      table('名冊找不到（rawdata 有紀錄，但統計檔沒有這個人）', ['rawdata', '卡號', '姓名', '導師', '座談日期'], r.notInRoster.map((x) => `<tr><td>${esc(T.CAT_LABEL[x.cat] || x.cat)}</td><td>${esc(x.card)}</td><td>${esc(x.name)}</td><td>${esc(x.tutor)}</td><td>${mmdd(x.date)}</td></tr>`)),
    ].join('');
  }

  $('tf-month').addEventListener('change', recompute);

  $('tf-download').addEventListener('click', async () => {
    const r = state.report;
    if (!r) return;
    if (!r.writes.length && !confirm('本月沒有需要填入的空白格，仍要下載嗎？')) return;
    const btn = $('tf-download');
    btn.disabled = true;
    try {
      const out = await F.patchXlsx(state.book.buffer, r.writes, { JSZip: window.JSZip, DOMParser, XMLSerializer });
      const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url;
      const tag = r.months.length === 1 ? `${Number(r.months[0].slice(0, 4)) - 1911}${r.months[0].slice(5, 7)}` : 'rawdata';
      a.download = `${state.book.name.replace(/\.xlsx$/i, '')}_${tag}已填入.xlsx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error(err);
      alert(`產生檔案失敗：${err.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  App.registerTab('tutor-fill', {
    reset() { state = emptyState(); $('tf-month').value = ''; recompute(); },
  });
  renderStatus();
})();
