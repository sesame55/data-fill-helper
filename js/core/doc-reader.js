/*
 * 讀取舊版 Word（.doc，Word 97-2003）的文字與表格（純前端，不需 Word）
 * 依 MS-DOC 規格：CFB 容器 → FIB → Clx 片段表 → 文字。
 * 表格儲存格以 \x07 結尾，每列最後多一個 \x07（列結束標記）。
 * 需要 SheetJS 的 XLSX.CFB（瀏覽器）或傳入 cfb 模組（Node 測試）。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory;
  else root.DocReader = factory(root.XLSX && root.XLSX.CFB);
})(typeof self !== 'undefined' ? self : this, (CFB) => {
  function getStream(cfb, name) {
    const entry = CFB.find(cfb, name);
    if (!entry || !entry.content) return null;
    const c = entry.content;
    return c instanceof Uint8Array ? c : new Uint8Array(c);
  }

  /** 取出 .doc 主文字（含表格標記字元） */
  function extractText(data) {
    const cfb = CFB.read(data instanceof Uint8Array ? data : new Uint8Array(data), { type: 'array' });
    const wd = getStream(cfb, 'WordDocument');
    if (!wd) throw new Error('不是 Word 97-2003 (.doc) 檔');
    const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    if (dv.getUint16(0, true) !== 0xA5EC) throw new Error('Word 檔頭不正確');
    const flags = dv.getUint16(0x0A, true);
    if (flags & 0x0100) throw new Error('Word 檔有加密，請先移除密碼');
    const table = getStream(cfb, (flags & 0x0200) ? '1Table' : '0Table');
    const fcClx = dv.getUint32(0x01A2, true);
    const lcbClx = dv.getUint32(0x01A6, true);
    const tv = new DataView(table.buffer, table.byteOffset, table.byteLength);

    // 跳過 Prc，找到 Pcdt
    let p = fcClx;
    const end = fcClx + lcbClx;
    while (p < end && table[p] === 0x01) p += 3 + tv.getUint16(p + 1, true);
    if (table[p] !== 0x02) throw new Error('Word 檔結構無法辨識');
    const lcb = tv.getUint32(p + 1, true);
    const plc = p + 5;
    const n = (lcb - 4) / 12;

    const decUtf16 = new TextDecoder('utf-16le');
    const decAnsi = new TextDecoder('windows-1252');
    let text = '';
    for (let i = 0; i < n; i += 1) {
      const cpStart = tv.getUint32(plc + i * 4, true);
      const cpEnd = tv.getUint32(plc + (i + 1) * 4, true);
      const pcd = plc + (n + 1) * 4 + i * 8;
      const fcRaw = tv.getUint32(pcd + 2, true);
      const len = cpEnd - cpStart;
      if (fcRaw & 0x40000000) {
        const fc = (fcRaw & 0x3FFFFFFF) / 2;
        text += decAnsi.decode(wd.subarray(fc, fc + len));
      } else {
        text += decUtf16.decode(wd.subarray(fcRaw, fcRaw + len * 2));
      }
    }
    return text;
  }

  const splitParas = (s) => s.split('\r').map((x) => x.trim()).filter(Boolean);

  /**
   * 解析成 { paragraphs: [文字], table: [[cell, ...], ...] }（取文件中的表格區）
   * 欄數以第一列（表頭，各格皆有文字）判斷，之後每 欄數+1 個片段為一列（最後一個是列結束標記）。
   */
  function parse(data) {
    // 保留 \x07（儲存格）與 \r（段落）；\x0b 換行轉 \n；其他控制字元移除
    const text = extractText(data).replace(/\x0b/g, '\n').replace(/[\x00-\x06\x08\x0c\x0e-\x1f]/g, '');
    const tokens = text.split('\x07');
    if (tokens.length < 3) return { paragraphs: splitParas(text), table: [] };

    const pre = tokens[0];
    const cut = pre.lastIndexOf('\r');
    const paragraphs = splitParas(pre.slice(0, cut + 1));
    const cells = [pre.slice(cut + 1), ...tokens.slice(1, -1)].map((c) => c.replace(/\r/g, '\n').trim());
    paragraphs.push(...splitParas(tokens[tokens.length - 1]));

    const ncols = cells.indexOf('');
    if (ncols <= 0) return { paragraphs, table: [] };
    const table = [];
    for (let i = 0; i + ncols <= cells.length; i += ncols + 1) table.push(cells.slice(i, i + ncols));
    return { paragraphs, table };
  }

  return { extractText, parse };
});
