/*
 * 手寫文字辨識：PaddleOCR PP-OCRv4 辨識模型（onnxruntime-web，於本機瀏覽器執行，不外傳）
 * 模型與字典放在 lib/paddle，執行環境放在 lib/onnx。
 * PaddleRec.recognize(canvas) → { text, conf }   （輸入為已切好的單一格子影像）
 */
(() => {
  let session = null;
  let keys = null;
  let loading = null;

  async function init() {
    if (session) return;
    if (loading) { await loading; return; }
    loading = (async () => {
      ort.env.wasm.wasmPaths = 'lib/onnx/';
      ort.env.wasm.numThreads = 1;   // GitHub Pages 無跨來源隔離，不能用多執行緒
      ort.env.wasm.proxy = false;
      const dict = await (await fetch('lib/paddle/ppocr_keys_v1.txt')).text();
      keys = ['blank', ...dict.split(/\r?\n/).filter((x, i, a) => i < a.length - 1 || x), ' '];
      session = await ort.InferenceSession.create('lib/paddle/ch_PP-OCRv4_rec_infer.onnx', { executionProviders: ['wasm'] });
    })();
    await loading;
  }

  /** 縮放到高 48（等比例），BGR、(x/255-0.5)/0.5 正規化，CTC 貪婪解碼 */
  async function recognize(src) {
    await init();
    const H = 48;
    const W = Math.max(16, Math.min(640, Math.round(src.width * H / Math.max(1, src.height))));
    const c = document.createElement('canvas');
    c.width = W; c.height = H;
    const cx = c.getContext('2d');
    cx.fillStyle = '#fff'; cx.fillRect(0, 0, W, H);
    cx.drawImage(src, 0, 0, W, H);
    const px = cx.getImageData(0, 0, W, H).data;
    const data = new Float32Array(3 * H * W);
    for (let i = 0; i < H * W; i += 1) {
      data[i] = (px[i * 4 + 2] / 255 - 0.5) / 0.5;           // B
      data[H * W + i] = (px[i * 4 + 1] / 255 - 0.5) / 0.5;   // G
      data[2 * H * W + i] = (px[i * 4] / 255 - 0.5) / 0.5;   // R
    }
    const res = await session.run({ [session.inputNames[0]]: new ort.Tensor('float32', data, [1, 3, H, W]) });
    const o = res[session.outputNames[0]];
    const [, T, C] = o.dims;
    let text = ''; let last = 0; let conf = 0; let n = 0;
    for (let t = 0; t < T; t += 1) {
      let bi = 0; let bv = -1;
      for (let k = 0; k < C; k += 1) { const v = o.data[t * C + k]; if (v > bv) { bv = v; bi = k; } }
      if (bi !== 0 && bi !== last) { text += keys[bi] || ''; conf += bv; n += 1; }
      last = bi;
    }
    return { text, conf: n ? conf / n : 0 };
  }

  window.PaddleRec = { init, recognize };
})();
