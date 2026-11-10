document.getElementById('go').addEventListener('click', async () => {
  const pw = document.getElementById('pw').value.trim().toLowerCase();   // 密碼不分大小寫
  if (!pw) return;
  const salt = [...crypto.getRandomValues(new Uint8Array(16))]
    .map((b) => b.toString(16).padStart(2, '0')).join('');
  const hash = await Auth.hashPassword(pw, salt);
  document.getElementById('out').textContent =
    `const SALT = '${salt}';\nconst HASH = '${hash}';`;
});
