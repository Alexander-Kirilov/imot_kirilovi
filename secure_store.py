"""Криптиране — репото е публично, затова данните и dashboard-ът се пазят криптирани.

Паролата идва от secret DASHBOARD_PASSWORD. С нея:
- dashboard-ът (docs/index.html) се криптира: страницата иска парола и се
  разкриптира в браузъра (AES-256-GCM, ключ от PBKDF2-SHA256);
- историите (*.parquet) и логът се пазят в репото само като *.enc.

Без парола скриптът спира, за да не запише нищо некриптирано.
При смяна на паролата: новата в DASHBOARD_PASSWORD, старата в
DASHBOARD_PASSWORD_OLD за едно пускане (данните се прекриптират), после
DASHBOARD_PASSWORD_OLD може да се изтрие.

Локално разкриптиране (напр. на лога): python secure_store.py imot_scraper.log.enc
"""
import base64
import getpass
import hashlib
import io
import os
import sys
from functools import lru_cache
from pathlib import Path

import pandas as pd
from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

PASSWORD_ENV = "DASHBOARD_PASSWORD"
OLD_PASSWORD_ENV = "DASHBOARD_PASSWORD_OLD"

MAGIC = b"IMK1"  # начало на всеки криптиран файл
DATA_SALT = b"imot_kirilovi/data/v1"
# Фиксирана сол за страницата, за да важи "запомни ме" и след обновяване на сайта
PAGE_SALT = b"imot_kirilovi/page/v1"
ITERATIONS = 250_000


def _password(env=PASSWORD_ENV):
    return os.environ.get(env, "")


def require_password():
    if not _password():
        raise SystemExit(f"{PASSWORD_ENV} не е зададена — спирам, за да не запиша "
                         f"некриптирани данни в публичното репо.")


@lru_cache(maxsize=None)
def _key(password, salt):
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, ITERATIONS, dklen=32)


def encrypt_bytes(data):
    require_password()
    iv = os.urandom(12)
    return MAGIC + iv + AESGCM(_key(_password(), DATA_SALT)).encrypt(iv, data, None)


def decrypt_bytes(blob):
    if not blob.startswith(MAGIC):
        raise SystemExit("Непознат формат на криптиран файл")
    iv, ct = blob[len(MAGIC):len(MAGIC) + 12], blob[len(MAGIC) + 12:]
    for env in (PASSWORD_ENV, OLD_PASSWORD_ENV):
        if _password(env):
            try:
                return AESGCM(_key(_password(env), DATA_SALT)).decrypt(iv, ct, None)
            except InvalidTag:
                continue
    # SystemExit, а не Exception: иначе скриптът ще продължи с празна история
    # и ще презапише натрупаните данни
    raise SystemExit("Данните не се разкриптират с тази парола — спирам, без да пипам нищо.")


def enc_path(path):
    return Path(f"{path}.enc")


def exists(path):
    return enc_path(path).exists() or Path(path).exists()


def read_parquet(path):
    """Чете path.enc; ако го няма — стария некриптиран path (еднократна миграция)."""
    path = Path(path)
    if enc_path(path).exists():
        return pd.read_parquet(io.BytesIO(decrypt_bytes(enc_path(path).read_bytes())))
    if path.exists():
        return pd.read_parquet(path)
    return pd.DataFrame()


def write_parquet(df, path):
    """Записва само криптирано копие; некриптираното (ако има) се трие."""
    path = Path(path)
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    enc_path(path).write_bytes(encrypt_bytes(buf.getvalue()))
    if path.exists():
        path.unlink()


def restore_file(path):
    """path.enc → path (напр. лога в началото на run-а)."""
    if enc_path(path).exists():
        Path(path).write_bytes(decrypt_bytes(enc_path(path).read_bytes()))


def seal_file(path):
    """path → path.enc (самият path е в .gitignore и не се качва)."""
    if Path(path).exists():
        enc_path(path).write_bytes(encrypt_bytes(Path(path).read_bytes()))


# ── Страницата с паролата ─────────────────────────────────────────────────────
# Съдържа само криптирания dashboard; разкриптира се в браузъра с WebCrypto.
LOGIN_PAGE = """<!DOCTYPE html>
<html lang="bg">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Имоти сем. Кирилови</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>🏠</text></svg>">
<style>
  :root { --bg: #0f1117; --surface: #1a1d27; --border: #2a2d3a; --text: #e2e4f0;
          --muted: #7a7d9a; --accent: #4f9cf9; --red: #f43f5e; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { min-height: 100vh; display: flex; align-items: center; justify-content: center;
         padding: 16px; background: var(--bg); color: var(--text);
         font-family: 'IBM Plex Sans', system-ui, sans-serif; font-size: 14px; }
  form { width: 100%; max-width: 360px; padding: 28px; background: var(--surface);
         border: 1px solid var(--border); border-radius: 12px; }
  h1 { font-size: 20px; font-weight: 600; margin-bottom: 4px; }
  h1 span { color: var(--accent); }
  p { color: var(--muted); font-size: 12px; margin-bottom: 20px; }
  input[type=password] { width: 100%; padding: 10px 12px; margin-bottom: 12px;
         background: var(--bg); border: 1px solid var(--border); border-radius: 6px;
         color: var(--text); font-size: 14px; outline: none; }
  input[type=password]:focus { border-color: var(--accent); }
  label { display: flex; gap: 8px; align-items: center; color: var(--muted);
          font-size: 12px; margin-bottom: 16px; cursor: pointer; }
  button { width: 100%; padding: 10px; background: var(--accent); color: #fff; border: 0;
           border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; }
  button:disabled { opacity: .6; cursor: wait; }
  .err { color: var(--red); font-size: 12px; margin-top: 12px; min-height: 16px; }
</style>
</head>
<body>
<form id="login" hidden>
  <h1>🏠 Имоти · <span>сем. Кирилови</span></h1>
  <p>Достъпът е само за семейството.</p>
  <input type="password" id="pw" placeholder="Парола" autocomplete="current-password" autofocus required>
  <label><input type="checkbox" id="remember" checked> Запомни ме на това устройство</label>
  <button type="submit" id="go">Влез</button>
  <div class="err" id="err"></div>
</form>
<script>
const PAYLOAD = "__PAYLOAD__";
const SALT = "__SALT__";
const ITERATIONS = __ITERATIONS__;
const STORE_KEY = 'imot_kirilovi_key';

const fromB64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
const toB64 = bytes => btoa(String.fromCharCode(...bytes));

async function deriveKey(password) {
  const base = await crypto.subtle.importKey(
    'raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits(
    { name: 'PBKDF2', hash: 'SHA-256', salt: fromB64(SALT), iterations: ITERATIONS }, base, 256);
  return new Uint8Array(bits);
}

async function decryptPage(rawKey) {
  const key = await crypto.subtle.importKey('raw', rawKey, 'AES-GCM', false, ['decrypt']);
  const data = fromB64(PAYLOAD);
  const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: data.slice(0, 12) }, key, data.slice(12));
  return new TextDecoder().decode(plain);
}

function showDashboard(html) {
  document.open();
  document.write(html);
  document.close();
}

function storage(action, value) {
  try {
    if (action === 'get') return localStorage.getItem(STORE_KEY);
    if (action === 'set') localStorage.setItem(STORE_KEY, value);
    if (action === 'del') localStorage.removeItem(STORE_KEY);
  } catch (e) { /* частен прозорец / блокирано хранилище */ }
  return null;
}

(async function () {
  // Запомнен ключ от предишно влизане
  const saved = storage('get');
  if (saved) {
    try { return showDashboard(await decryptPage(fromB64(saved))); }
    catch (e) { storage('del'); }  // паролата е сменена
  }
  const form = document.getElementById('login');
  form.hidden = false;
  document.getElementById('pw').focus();
  form.addEventListener('submit', async e => {
    e.preventDefault();
    const btn = document.getElementById('go');
    const err = document.getElementById('err');
    btn.disabled = true;
    err.textContent = '';
    try {
      const rawKey = await deriveKey(document.getElementById('pw').value);
      const html = await decryptPage(rawKey);
      if (document.getElementById('remember').checked) storage('set', toB64(rawKey));
      showDashboard(html);
    } catch (ex) {
      err.textContent = window.crypto && crypto.subtle ? 'Грешна парола.' : 'Браузърът не поддържа разкриптиране.';
      btn.disabled = false;
    }
  });
})();
</script>
</body>
</html>
"""


def encrypt_page(html):
    """Dashboard HTML → страница с поле за парола, която го разкриптира в браузъра."""
    require_password()
    key = _key(_password(), PAGE_SALT)
    iv = os.urandom(12)
    payload = base64.b64encode(iv + AESGCM(key).encrypt(iv, html.encode("utf-8"), None)).decode()
    return (LOGIN_PAGE
            .replace("__PAYLOAD__", payload)
            .replace("__SALT__", base64.b64encode(PAGE_SALT).decode())
            .replace("__ITERATIONS__", str(ITERATIONS)))


if __name__ == "__main__":
    # Локално разкриптиране: python secure_store.py <файл.enc> [<файл.enc> …]
    if not _password():
        os.environ[PASSWORD_ENV] = getpass.getpass("Парола: ")
    for name in sys.argv[1:]:
        src = Path(name)
        dst = src.with_suffix("") if src.suffix == ".enc" else Path(f"{src}.dec")
        dst.write_bytes(decrypt_bytes(src.read_bytes()))
        print(f"{src} → {dst}")
