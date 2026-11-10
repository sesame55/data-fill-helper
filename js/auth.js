/*
 * 簡易登入（純前端）
 * - 密碼不以明文存放，只存 PBKDF2-SHA256 雜湊值。
 * - 登入狀態只存在 sessionStorage：關閉分頁即失效。
 * - 修改密碼：開啟 tools/make-hash.html 產生新的 SALT / HASH，貼到下方。
 *
 * 注意：GitHub Pages 是公開網站，前端登入只能擋住一般人誤用，
 *       真正的資料保護在於「資料從不離開使用者電腦」。
 */
const Auth = (() => {
  // ★ 登入開關：開發測試時 false；上傳 GitHub 前改回 true
  const ENABLED = true;

  const SALT = '448ffd61c86e55d8a8e9fcd0f8cdf507';
  const HASH = '5424ba4f0b3eb1175898ef922d5dd1272e27e1f7cfb3689606d68672bd9e93cd';
  const ITERATIONS = 150000;
  const SESSION_KEY = 'auth-ok';
  const IDLE_MINUTES = 15;            // 閒置多久自動登出
  const MAX_ATTEMPTS = 5;             // 連續錯誤次數上限
  const LOCK_SECONDS = 60;            // 超過上限鎖定秒數

  let failCount = 0;
  let lockedUntil = 0;

  const hexToBytes = (hex) => new Uint8Array(hex.match(/../g).map((h) => parseInt(h, 16)));
  const bytesToHex = (buf) => [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');

  async function hashPassword(password, saltHex = SALT) {
    const key = await crypto.subtle.importKey(
      'raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveBits']);
    const bits = await crypto.subtle.deriveBits(
      { name: 'PBKDF2', hash: 'SHA-256', salt: hexToBytes(saltHex), iterations: ITERATIONS },
      key, 256);
    return bytesToHex(bits);
  }

  async function login(password) {
    const now = Date.now();
    if (now < lockedUntil) {
      throw new Error(`嘗試次數過多，請 ${Math.ceil((lockedUntil - now) / 1000)} 秒後再試`);
    }
    // 不分大小寫：一律轉小寫後比對（產生雜湊時也轉小寫）
    const ok = (await hashPassword(password.trim().toLowerCase())) === HASH;
    if (!ok) {
      failCount += 1;
      if (failCount >= MAX_ATTEMPTS) {
        lockedUntil = now + LOCK_SECONDS * 1000;
        failCount = 0;
      }
      throw new Error('密碼錯誤');
    }
    failCount = 0;
    sessionStorage.setItem(SESSION_KEY, '1');
  }

  function isLoggedIn() {
    return sessionStorage.getItem(SESSION_KEY) === '1';
  }

  function logout() {
    sessionStorage.removeItem(SESSION_KEY);
  }

  return { ENABLED, login, logout, isLoggedIn, hashPassword, IDLE_MINUTES };
})();
