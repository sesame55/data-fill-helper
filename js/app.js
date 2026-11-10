/*
 * 主程式：登入流程、分頁切換、閒置自動登出、清空資料。
 *
 * 新增分頁的方式：
 *   1. index.html 的 #tabs 加一顆 <button class="tab" data-tab="xxx">
 *   2. 加一個 <section id="tab-xxx" class="tab-panel">
 *   3. 在 js/tabs/xxx.js 呼叫 App.registerTab('xxx', { reset() {...} })
 */
const App = (() => {
  const tabModules = {};   // 各分頁自己的資料（只存在記憶體）
  let idleTimer = null;
  let idleDeadline = 0;

  const $ = (sel) => document.querySelector(sel);

  function registerTab(name, module) {
    tabModules[name] = module;
  }

  function showTab(name) {
    document.querySelectorAll('.tab').forEach((t) =>
      t.classList.toggle('active', t.dataset.tab === name));
    document.querySelectorAll('.tab-panel').forEach((p) =>
      p.classList.toggle('active', p.id === `tab-${name}`));
  }

  function clearAllData() {
    Object.values(tabModules).forEach((m) => m.reset && m.reset());
  }

  function showApp() {
    $('#login-view').hidden = true;
    $('#app-view').hidden = false;
    $('#idle-minutes').textContent = Auth.IDLE_MINUTES;
    resetIdle();
  }

  function showLogin() {
    clearAllData();
    Auth.logout();
    $('#app-view').hidden = true;
    $('#login-view').hidden = false;
    $('#password').value = '';
    $('#password').focus();
    clearInterval(idleTimer);
  }

  // ---- 閒置自動登出 ----
  function resetIdle() {
    idleDeadline = Date.now() + Auth.IDLE_MINUTES * 60 * 1000;
  }
  function startIdleWatch() {
    clearInterval(idleTimer);
    idleTimer = setInterval(() => {
      const left = idleDeadline - Date.now();
      if (left <= 0) {
        showLogin();
        alert('閒置過久，已自動登出並清除資料。');
        return;
      }
      const m = Math.floor(left / 60000);
      const s = Math.floor((left % 60000) / 1000);
      $('#idle-timer').textContent = `自動登出 ${m}:${String(s).padStart(2, '0')}`;
    }, 1000);
  }

  function init() {
    $('#login-form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const err = $('#login-error');
      err.hidden = true;
      try {
        await Auth.login($('#password').value);
        $('#password').value = '';
        showApp();
        startIdleWatch();
      } catch (ex) {
        err.textContent = ex.message;
        err.hidden = false;
      }
    });

    $('#tabs').addEventListener('click', (e) => {
      const btn = e.target.closest('.tab');
      if (btn) showTab(btn.dataset.tab);
    });
    document.querySelectorAll('[data-goto]').forEach((el) =>
      el.addEventListener('click', () => showTab(el.dataset.goto)));

    // 分頁內的子分頁（例如 教學住診 / 教學門診）
    document.querySelectorAll('.subtabs').forEach((bar) => bar.addEventListener('click', (e) => {
      const btn = e.target.closest('.subtab[data-sub]');
      if (!btn || btn.disabled) return;
      const panel = bar.closest('.tab-panel');
      bar.querySelectorAll('.subtab').forEach((b) => b.classList.toggle('active', b === btn));
      panel.querySelectorAll('[data-subpanel]').forEach((p) => { p.hidden = p.dataset.subpanel !== btn.dataset.sub; });
    }));

    $('#logout-btn').addEventListener('click', showLogin);
    $('#clear-btn').addEventListener('click', () => {
      if (confirm('確定清除所有分頁已匯入的資料？')) clearAllData();
    });

    ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'].forEach((ev) =>
      document.addEventListener(ev, resetIdle, { passive: true }));

    // 登入關閉時：直接進入，不顯示登出與閒置倒數
    if (!Auth.ENABLED) {
      $('#logout-btn').hidden = true;
      $('#idle-timer').hidden = true;
      $('#idle-note').hidden = true;
      $('#login-view').hidden = true;
      $('#app-view').hidden = false;
      return;
    }

    if (Auth.isLoggedIn()) {
      showApp();
      startIdleWatch();
    } else {
      $('#login-view').hidden = false;
      $('#password').focus();
    }
  }

  document.addEventListener('DOMContentLoaded', init);
  return { registerTab, showTab, clearAllData };
})();
