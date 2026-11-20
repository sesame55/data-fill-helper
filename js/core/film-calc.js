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
      return { ...e, votes: tally.get(e.n) || 0, dept: depts.join('、'), flag, note, ...styleOf(essayData).get(e.n) };
    });
    const style = { short: 0, levels: { 無法判斷: 0, 低: 0, 中: 0, 高: 0 }, copied: 0 };
    rows.forEach((r) => { if (r.short) style.short++; style.levels[r.level]++; if (r.copy) style.copied++; });
    return { rows, detail, stats, match, style, keywords: kw };
  }

  /* ======================================================================
   * 字數與 AI 風格／複製參考（僅供參考）
   * 字數填色與複製比對依「(完成版)心得字數_AI風格_複製標示.xlsx」歸納；AI 風格為自訂規則（見 judgeAI）
   * ====================================================================== */
  const SHORT_HAN = 100;        // 中文字數 < 100 → 心得格填綠色
  const TOO_SHORT_HAN = 60;     // 中文字數 < 60 → 字數太少、無法判斷
  const SHORT_CAP_HAN = 150;    // 不足 150 字：資訊太少，最高只判「中」

  /*
   * AI 風格判斷（自訂規則，僅供參考）
   *   加分：AI 常見句型（成對句型、說教收尾、總結語、文藝套語）、段落工整、書面標點、資料性內容、篇幅長
   *   減分：口語用詞、非標準標點、個人經驗
   *   等級：不足 60 字＝無法判斷；分數 ≥ 7＝高、≥ 4＝中、其餘＝低
   */
  const AI_PATTERNS = [
    ['不只是…更是', /不(只|僅|單)(是|在於)?[^。！？\n]{1,40}(更是|而是|也是|更在於)/],
    ['不是…而是', /不是[^。！？\n]{1,25}[，,]\s*而是/],
    ['不僅…更', /不僅[^。！？\n]{1,30}[，,]?\s*(更|還|也)/],
    ['電影提醒／告訴我們', /(電影|影片|這部片|本片|它|故事)[^。！？\n]{0,8}(提醒|告訴|教會|讓)(了)?(我們|觀眾)/],
    ['總結語', /(總而言之|總結來說|總的來說|整體而言|整體來說|綜上所述|總體而言)/],
    ['轉折連接詞', /(此外|與此同時|另一方面|更重要的是|換句話說|事實上|值得一提的是|(^|[。！？\n])\s*同時，)/],
    ['說教式收尾', /(讓我(明白|學到|學習到|體會到|意識到|了解到)|我學習到|未來無論|期許自己|希望自己能|成為一個更)/],
    ['首先／其次／最後', /首先[^。]*。[\s\S]*(其次|再者)/],
    ['在…的過程中', /在[^，。！？\n]{1,20}的過程中/],
    ['值得深思', /(值得(我們)?(深思|省思|反思|細細品味|一看再看)|引人深思|發人深省)/],
    ['文藝套語', /(靈魂|救贖|共鳴|交織|淋漓盡致|絲絲入扣|娓娓道來|耐人尋味|不落俗套|恰到好處|細膩地|巧妙地|溫柔地)/],
    ['真正的…', /真正(的)?(意義|價值|勇氣|自由|幸福|的愛|重要的)/],
  ];
  const COLLOQUIAL_WORDS = /(老實說|說真的|真的很|真的超|超級|超好看|好好看|有點|蠻|挺好|啦|吧|耶|喔|欸|哈哈|嗯|沒想到|好險|還好|看不懂|推薦|很好看|不錯看|幹話|鳥事)/g;
  const ODD_PUNCT = /[一-鿿][,.;:!?]|[,.;:!?][一-鿿]|[~～]|\.{2,}|[!！]{2,}|[?？]{2,}|，，|。。|\s[，。、]|ㄧ/g;
  // 評論腔用語：出現 3 種以上 +1、5 種以上 +2
  const REVIEW_TERMS = ['載體', '解構', '催化劑', '視角', '核心', '探討', '呈現', '描繪', '詮釋', '隱喻', '象徵', '意象', '張力',
    '敘事', '層次', '基調', '筆觸', '鏡頭語言', '弧線', '投射', '映照', '折射', '昇華', '療癒', '框架', '議題'];
  const PERSONAL = /(我(的)?(女兒|兒子|孩子|小孩|先生|太太|老公|老婆|媽媽|爸爸|父母|家人|阿嬤|阿公|同事|病人)|身為(一[名位個])?(護理師|醫師|醫生|父親|母親|媽媽|爸爸|醫療人員)|在(病房|臨床|門診|急診)(工作|服務|中|裡)|我自己|我以前|我曾經|小時候)/;
  const FORMAL_PUNCT = /——|；|：「/;
  const FACTUAL = /《[^》]*[A-Za-z][^》]*》|（[A-Za-z][^）]*）|City\s?Pop|Plastic\s?Love|竹內瑪莉亞|妮可葉特|克雷比茨|蘇菲[·・]?蘿伊絲|米蘭[·・]?赫姆斯/;

  /** 段落長度是否很平均（≥3 段、長度變異係數 < 0.35） */
  function evenParagraphs(t) {
    const lens = t.split('\n').map(hanCount).filter((n) => n >= 15);
    if (lens.length < 3) return false;
    const avg = lens.reduce((s, n) => s + n, 0) / lens.length;
    const sd = Math.sqrt(lens.reduce((s, n) => s + (n - avg) ** 2, 0) / lens.length);
    return sd / avg < 0.35;
  }

  /** 回傳 { level, reasons[] } */
  function judgeAI(t, han) {
    if (han < TOO_SHORT_HAN) return { level: '無法判斷', reasons: ['字數太少'] };
    let score = 0;
    const reasons = [];
    const hits = AI_PATTERNS.filter(([, re]) => re.test(t)).map(([name]) => name);
    if (hits.length) { score += Math.min(hits.length, 5); reasons.push(`AI常見句型：${hits.join('、')}`); }
    const paras = t.split('\n').filter((l) => hanCount(l) >= 15).length;
    if (evenParagraphs(t)) { score += 2; reasons.push(`分段工整（${paras} 段、長度平均）`); }
    const lines = t.split('\n').map((l) => l.trim()).filter(Boolean);
    const titled = lines.length >= 2 && hanCount(lines[0]) <= 40 && !/[。！？]$/.test(lines[0]) && hanCount(lines.slice(1).join('')) >= 100;
    if (titled || /(^|\n)\s*([一二三四五六七八九十]、|\d+[.、]\s*\S|[*#]{1,3}\s*\S|【)/.test(t)) { score += 2; reasons.push('有標題、小標題或條列'); }
    const terms = REVIEW_TERMS.filter((w) => t.includes(w));
    if (terms.length >= 3) { score += terms.length >= 5 ? 2 : 1; reasons.push(`評論腔用語：${terms.slice(0, 5).join('、')}`); }
    const odd = (t.match(ODD_PUNCT) || []).length;
    if (FORMAL_PUNCT.test(t) && !odd) { score += 1; reasons.push('書面標點（破折號、分號等）'); }
    if (FACTUAL.test(t)) { score += 1; reasons.push('含英文片名、導演演員全名或歌曲等資料'); }
    if (han >= 400) { score += 1; reasons.push('篇幅長'); }
    // 口語與非標準標點：長文只偶爾出現一次扣 1，否則扣 2
    const casual = [...new Set(t.match(COLLOQUIAL_WORDS) || [])];
    if (casual.length) { score -= casual.length === 1 && han >= 400 ? 1 : 2; reasons.push(`口語用詞（${casual.slice(0, 3).map((w) => `「${w}」`).join('')}）`); }
    if (odd) { score -= odd === 1 && han >= 400 ? 1 : 2; reasons.push('非標準標點'); }
    const personal = PERSONAL.exec(t);
    if (personal) { score -= 1; reasons.push(`個人經驗（如「${personal[0]}」）`); }
    let level = score >= 7 ? '高' : score >= 4 ? '中' : '低';
    if (level === '高' && han < SHORT_CAP_HAN) { level = '中'; reasons.push('字數少，最高判中'); }
    if (!reasons.length) reasons.push('未見明顯AI特徵');
    return { level, reasons, score };
  }
  const COPY_MIN_BLOCK = 8;     // 相同片段至少 8 字（去空白標點後）
  const COPY_RATIO = 0.5;       // 相同片段合計 ≥ 較短一篇的 50%
  const COPY_MIN_LEN = 20;      // 去空白標點後不足 20 字的不比對
  const COPY_SHOW = 4;          // 最多列 4 篇（編號小的優先）

  const hanCount = (t) => (t.match(/[一-鿿]/g) || []).length;
  const copyNorm = (t) => t.replace(/[^\p{L}\p{N}]/gu, '');

  /** Python difflib.SequenceMatcher(None, a, b, autojunk=False).get_matching_blocks() 的移植（結果與 Python 相同） */
  function matchingBlocks(a, b) {
    const b2j = new Map();
    for (let j = 0; j < b.length; j++) { const c = b[j]; if (!b2j.has(c)) b2j.set(c, []); b2j.get(c).push(j); }
    const longest = (alo, ahi, blo, bhi) => {
      let besti = alo; let bestj = blo; let bestsize = 0;
      let j2len = new Map();
      for (let i = alo; i < ahi; i++) {
        const next = new Map();
        for (const j of b2j.get(a[i]) || []) {
          if (j < blo) continue;
          if (j >= bhi) break;
          const k = (j2len.get(j - 1) || 0) + 1;
          next.set(j, k);
          if (k > bestsize) { besti = i - k + 1; bestj = j - k + 1; bestsize = k; }
        }
        j2len = next;
      }
      while (besti > alo && bestj > blo && a[besti - 1] === b[bestj - 1]) { besti--; bestj--; bestsize++; }
      while (besti + bestsize < ahi && bestj + bestsize < bhi && a[besti + bestsize] === b[bestj + bestsize]) bestsize++;
      return [besti, bestj, bestsize];
    };
    const queue = [[0, a.length, 0, b.length]];
    const found = [];
    while (queue.length) {
      const [alo, ahi, blo, bhi] = queue.pop();
      const [i, j, k] = longest(alo, ahi, blo, bhi);
      if (k) {
        found.push([i, j, k]);
        if (alo < i && blo < j) queue.push([alo, i, blo, j]);
        if (i + k < ahi && j + k < bhi) queue.push([i + k, ahi, j + k, bhi]);
      }
    }
    found.sort((x, y) => x[0] - y[0] || x[1] - y[1] || x[2] - y[2]);
    const out = [];
    let [i1, j1, k1] = [0, 0, 0];
    for (const [i2, j2, k2] of found) {
      if (i1 + k1 === i2 && j1 + k1 === j2) k1 += k2;
      else { if (k1) out.push(k1); [i1, j1, k1] = [i2, j2, k2]; }
    }
    if (k1) out.push(k1);
    return out; // 只需要各區塊長度
  }

  /** 心得之間的相同段落：回傳 Map(編號 → [對方編號…]) */
  function copyPartners(essays) {
    const items = essays.map((e) => ({ n: e.n, t: copyNorm(e.text) })).filter((x) => x.t.length >= COPY_MIN_LEN);
    // 先用 8 字片段建索引，只比對至少有一段相同的配對
    const index = new Map();
    items.forEach((x, idx) => {
      const seen = new Set();
      for (let i = 0; i + COPY_MIN_BLOCK <= x.t.length; i++) {
        const g = x.t.slice(i, i + COPY_MIN_BLOCK);
        if (seen.has(g)) continue;
        seen.add(g);
        if (!index.has(g)) index.set(g, []);
        index.get(g).push(idx);
      }
    });
    const cand = new Set();
    for (const list of index.values()) {
      if (list.length < 2) continue;
      for (let p = 0; p < list.length; p++) for (let q = p + 1; q < list.length; q++) cand.add(list[p] * 100000 + list[q]);
    }
    const partners = new Map(items.map((x) => [x.n, []]));
    for (const key of cand) {
      const A = items[Math.floor(key / 100000)];
      const B = items[key % 100000];
      const same = matchingBlocks(A.t, B.t).filter((k) => k >= COPY_MIN_BLOCK).reduce((s, k) => s + k, 0);
      if (same / Math.min(A.t.length, B.t.length) >= COPY_RATIO) { partners.get(A.n).push(B.n); partners.get(B.n).push(A.n); }
    }
    for (const list of partners.values()) list.sort((x, y) => x - y);
    return partners;
  }

  /** 每篇心得的字數、AI 風格等級與判斷依據、複製參考 */
  function styleCheck(essays) {
    const partners = copyPartners(essays);
    const out = new Map();
    for (const e of essays) {
      const t = e.text;
      const han = hanCount(t);
      const others = (partners.get(e.n) || []).slice(0, COPY_SHOW).map((n) => String(n).padStart(4, '0'));
      const simNote = others.length ? `與編號 ${others.join('、')} 內容高度相似` : '';
      const { level, reasons, score } = judgeAI(t, han);
      if (simNote) { if (reasons[0] === '未見明顯AI特徵') reasons.shift(); reasons.push(simNote); }
      const pdf = others.length && /[一-鿿] [一-鿿]/.test(t);
      const copy = others.length ? `與編號 ${others.join('、')} 有相同段落${pdf ? '，且有PDF斷行留下的空格，可能複製自同一份講義' : ''}` : '';
      out.set(e.n, { han, short: han < SHORT_HAN, level, score, reason: reasons.join('；'), copy });
    }
    return out;
  }
  // 只跟心得有關，算一次就快取在心得資料上（改關鍵字不必重算）
  const styleOf = (essayData) => essayData.style || (essayData.style = styleCheck(essayData.essays));

  return { VOTE_SHEET, DEFAULT_KEYWORDS, SHORT_HAN, detectType, parseEssays, parseBallot, parseVotes, analyze, normName, styleCheck, matchingBlocks };
});
