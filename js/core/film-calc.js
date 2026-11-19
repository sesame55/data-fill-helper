/*
 * 全人電影：心得張貼 ＋ 票選資料 → 心得票數表（純計算，不碰 DOM）
 *
 * 計票規則：
 *   - 一律解析「符合票選資格」的原始填答（不用 Excel 已切好的三欄，那三欄常切錯）。
 *   - 全形數字轉半形；去掉「1.」「2，」這類序號前綴；任何非數字都當分隔。
 *   - 連在一起的長數字：每 4 位且都以 0 開頭 → 4 位切；否則長度是 3 的倍數 → 3 位切；都不是 → 不計。
 *   - 只算心得檔中真的有心得的編號；同一張票重複的號碼只算一次；每張票最多 3 票（取前 3 個有效號碼）。
 * 部門對照：
 *   - 心得作者姓名（去空白、NFKC）對「符合票選資格」的姓名 → 部門；找不到再查其他名單工作表（如 Sheet1）。
 *   - 部門含關鍵字（預設 聯合、台南、臺南）者標「是」。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FilmCalc = factory();
})(typeof self !== 'undefined' ? self : this, () => {
  const VOTE_SHEET = '符合票選資格';
  const DEFAULT_KEYWORDS = ['聯合', '台南', '臺南'];

  const str = (v) => (v == null ? '' : String(v));
  const normName = (s) => str(s).normalize('NFKC').replace(/[\s 　]/g, '');
  const cleanText = (s) => str(s).trim(); // 心得內容保持原樣，只去頭尾空白
  const cleanName = (s) => str(s).replace(/[\s 　]+/g, ' ').trim();

  /* ---------- 檔案辨識 ---------- */
  // sheets: [{ name, rows: [[...], ...] }]
  function detectType(sheets) {
    if (sheets.some((s) => s.name.trim() === VOTE_SHEET)) return 'votes';
    const essaySheets = sheets.filter((s) => /^\d+\s*-\s*\d+$/.test(s.name.trim()));
    if (essaySheets.length && essaySheets.length >= sheets.length / 2) return 'essays';
    return null;
  }

  /* ---------- 心得 ---------- */
  // 每個工作表：A 編號、B 心得、C 作者。整頁沒有任何作者的工作表（例如「001-100」說明頁）略過。
  function parseEssays(sheets) {
    const essays = [];
    const skipped = [];
    for (const s of sheets) {
      const rows = s.rows.filter((r) => /^\d{1,5}$/.test(str(r[0]).trim()));
      if (!rows.some((r) => cleanName(r[2]))) { if (rows.some((r) => cleanText(r[1]))) skipped.push(s.name); continue; }
      for (const r of rows) {
        const text = cleanText(r[1]);
        const author = cleanName(r[2]);
        if (!text && !author) continue;
        const n = Number(str(r[0]).trim());
        essays.push({ no: String(n).padStart(4, '0'), n, text, author, sheet: s.name });
      }
    }
    essays.sort((a, b) => a.n - b.n);
    return { essays, skipped };
  }

  /* ---------- 單張選票解析 ---------- */
  const FW = { '０': '0', '１': '1', '２': '2', '３': '3', '４': '4', '５': '5', '６': '6', '７': '7', '８': '8', '９': '9' };
  function parseBallot(raw, essaySet) {
    let s = str(raw).replace(/[０-９]/g, (c) => FW[c]);
    // 序號前綴：行首或分隔符後的 1~3，接標點，再接 3~4 位數字
    s = s.replace(/(^|[\s,，、;；])[123]\s*[.．、，,]\s*(?=\d{3,4}(?!\d))/g, '$1 ');
    const nums = [];
    const notes = [];
    for (const t of s.match(/\d+/g) || []) {
      if (t.length <= 4) { nums.push(Number(t)); continue; }
      let parts = null;
      if (t.length % 4 === 0 && [...Array(t.length / 4).keys()].every((i) => t[i * 4] === '0')) parts = t.match(/\d{4}/g);
      else if (t.length % 3 === 0) parts = t.match(/\d{3}/g);
      if (parts) { nums.push(...parts.map(Number)); notes.push(`${t} 切分為 ${parts.join('、')}`); } else notes.push(`${t} 無法切分`);
    }
    const valid = [];
    const invalid = [];
    const dup = [];
    for (const n of nums) {
      if (!essaySet.has(n)) invalid.push(n);
      else if (valid.includes(n)) dup.push(n);
      else valid.push(n);
    }
    const extra = valid.splice(3);
    if (invalid.length) notes.push(`無效號碼：${invalid.join('、')}`);
    if (dup.length) notes.push(`重複號碼只算一次：${dup.join('、')}`);
    if (extra.length) notes.push(`超過三票，不計：${extra.join('、')}`);
    return { valid, notes, dup: dup.length > 0, extra: extra.length > 0 };
  }

  /* ---------- 票選資料 ---------- */
  const findCol = (header, re) => header.findIndex((h) => re.test(str(h)));

  function parseVotes(sheets) {
    const ws = sheets.find((s) => s.name.trim() === VOTE_SHEET);
    const header = ws.rows[0] || [];
    const col = {
      acct: Math.max(0, findCol(header, /帳號|卡號/)),
      name: Math.max(1, findCol(header, /姓名/)),
      dept: Math.max(2, findCol(header, /部門|單位/)),
    };
    let ballot = findCol(header, /最喜歡|票/);
    if (ballot < 0) ballot = 5;
    const ballots = [];
    ws.rows.slice(1).forEach((r, i) => {
      if (![0, 1, 2, 3, 4, ballot].some((c) => str(r[c]).trim())) return;
      ballots.push({ row: i + 2, acct: str(r[col.acct]).trim(), name: cleanName(r[col.name]), dept: str(r[col.dept]).trim(), raw: str(r[ballot]) });
    });

    // 部門名單：主名單＝符合票選資格；備用＝其他看起來是「帳號,姓名,部門…」的工作表（例如 Sheet1）
    const primary = new Map();
    for (const b of ballots) {
      const k = normName(b.name);
      if (!k || !b.dept) continue;
      if (!primary.has(k)) primary.set(k, []);
      if (!primary.get(k).includes(b.dept)) primary.get(k).push(b.dept);
    }
    const backup = new Map();
    for (const s of sheets) {
      if (s === ws) continue;
      const h = s.rows[0] || [];
      let ni = findCol(h, /^姓名$/);
      let di = findCol(h, /^(部門|單位)$/);
      let start = 1;
      if (ni < 0 || di < 0) {
        // 無標題列：A 帳號、B 姓名、C 部門
        const looks = s.rows.slice(0, 5).filter((r) => /^[a-z]?\d{3,5}$/i.test(str(r[0]).trim()) && /\D/.test(str(r[2])) && str(r[1]).trim());
        if (looks.length < Math.min(3, s.rows.length)) continue;
        ni = 1; di = 2; start = 0;
      }
      for (const r of s.rows.slice(start)) {
        const k = normName(r[ni]);
        const d = str(r[di]).trim();
        if (k && d && !backup.has(k)) backup.set(k, { dept: d, sheet: s.name });
      }
    }
    return { ballots, primary, backup };
  }

  /* ---------- 彙整 ---------- */
  function analyze(essayData, voteData, keywords = DEFAULT_KEYWORDS) {
    const { essays } = essayData;
    const essaySet = new Set(essays.map((e) => e.n));
    const tally = new Map();
    const stats = { ballots: 0, validVotes: 0, noValid: 0, dupBallots: 0, extraBallots: 0 };
    const detail = voteData.ballots.map((b) => {
      const p = parseBallot(b.raw, essaySet);
      p.valid.forEach((n) => tally.set(n, (tally.get(n) || 0) + 1));
      stats.ballots++;
      stats.validVotes += p.valid.length;
      if (!p.valid.length) stats.noValid++;
      if (p.dup) stats.dupBallots++;
      if (p.extra) stats.extraBallots++;
      return { ...b, valid: p.valid, note: p.notes.join('；') };
    });

    const kw = keywords.map((k) => k.trim()).filter(Boolean);
    const hasKw = (d) => kw.some((k) => d.includes(k));
    const match = { primary: 0, backup: 0, notFound: 0, multi: 0, noAuthor: 0, flagged: 0 };
    const rows = essays.map((e) => {
      const k = normName(e.author);
      let depts = [];
      let note = '';
      if (!k) { note = '無作者姓名'; match.noAuthor++; }
      else if (voteData.primary.has(k)) {
        depts = voteData.primary.get(k); match.primary++;
        if (depts.length > 1) { note = '同名多筆，請人工確認'; match.multi++; }
      } else if (voteData.backup.has(k)) {
        const b = voteData.backup.get(k);
        depts = [b.dept]; note = `由 ${b.sheet} 對照`; match.backup++;
      } else { note = '02 查無此人'; match.notFound++; }
      const flag = depts.some(hasKw);
      if (flag) match.flagged++;
      return { ...e, votes: tally.get(e.n) || 0, dept: depts.join('、'), flag, note };
    });
    return { rows, detail, stats, match, keywords: kw };
  }

  return { VOTE_SHEET, DEFAULT_KEYWORDS, detectType, parseEssays, parseBallot, parseVotes, analyze, normName };
});
