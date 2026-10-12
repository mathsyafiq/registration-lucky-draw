// ============================================================
// Shared helpers for all pages
// ============================================================

// ── Configuration ────────────────────────────────────────────
// Paste your Apps Script Web App URL here (ends in /exec).
var DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbzo19hbZ_aMODI3UfR_LhBEOm-OzIo_umnegDtGmOoxc1JoJzvvv5-arMoIh7c3uALL/exec';

var App = (function () {
  // Another organiser can reuse these pages with their own Google Sheet
  // by adding ?api=<their web app URL> to the link.
  var params = new URLSearchParams(location.search);
  var override = params.get('api');
  if (override && !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(override)) override = null;
  var API_URL = override || DEFAULT_API_URL;

  function isConfigured() {
    return API_URL.indexOf('XXXX') === -1;
  }

  // Keep ?api= when moving between pages
  function link(page) {
    return override ? page + '?api=' + encodeURIComponent(override) : page;
  }

  function absoluteLink(page) {
    return new URL(link(page), location.href).href;
  }

  function get(query) {
    var url = API_URL + '?' + new URLSearchParams(query || {}).toString();
    return fetch(url).then(function (r) { return r.json(); });
  }

  // text/plain avoids a CORS preflight, which Apps Script cannot answer
  function post(body) {
    return fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function (r) { return r.json(); });
  }

  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  // ── Storage (wrapped: private mode can throw) ──────────────
  function storeKey(k) { return 'ld:' + API_URL + ':' + k; }
  function session(k, v) {
    try {
      if (v === undefined) return sessionStorage.getItem(storeKey(k));
      if (v === null) sessionStorage.removeItem(storeKey(k));
      else sessionStorage.setItem(storeKey(k), v);
    } catch (e) { return null; }
  }
  function local(k, v) {
    try {
      if (v === undefined) return localStorage.getItem(storeKey(k));
      if (v === null) localStorage.removeItem(storeKey(k));
      else localStorage.setItem(storeKey(k), v);
    } catch (e) { return null; }
  }

  // ── Toast ──────────────────────────────────────────────────
  var toastTimer;
  function toast(msg, type) {
    var el = document.getElementById('toast');
    if (!el) {
      el = document.createElement('div');
      el.id = 'toast';
      el.className = 'toast';
      document.body.appendChild(el);
    }
    el.textContent = msg;
    el.className = 'toast show ' + (type || '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.className = 'toast ' + (type || ''); }, 3500);
  }

  // ── PIN gate used by Admin and Lucky Draw pages ────────────
  // Calls onReady(pin, loginResponse) once a valid PIN is known.
  function pinGate(onReady) {
    var gate = document.getElementById('pinGate');
    var form = document.getElementById('pinForm');
    var input = document.getElementById('pinInput');
    var err = document.getElementById('pinError');
    var btn = form.querySelector('button');

    if (override) {
      var note = document.getElementById('pinBackend');
      if (note) note.textContent = 'Custom backend: …' + override.slice(-18, -5);
    }

    function tryPin(pin, silent) {
      btn.disabled = true;
      err.textContent = '';
      return post({ action: 'login', pin: pin }).then(function (res) {
        btn.disabled = false;
        if (res.success) {
          session('pin', pin);
          gate.classList.add('hidden');
          onReady(pin, res);
        } else {
          session('pin', null);
          gate.classList.remove('hidden');
          if (!silent) err.textContent = res.message || 'Wrong PIN.';
          input.focus();
        }
      }).catch(function () {
        btn.disabled = false;
        gate.classList.remove('hidden');
        err.textContent = 'Cannot reach the server. Check your connection.';
      });
    }

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (input.value.trim()) tryPin(input.value.trim(), false);
    });

    if (!isConfigured()) {
      err.textContent = 'API URL not set. Edit DEFAULT_API_URL in docs/assets/app.js.';
      return;
    }
    var saved = session('pin');
    if (saved) tryPin(saved, true);
    else input.focus();
  }

  function logout() {
    session('pin', null);
    location.reload();
  }

  function setupLinks() {
    document.querySelectorAll('a[data-page]').forEach(function (a) {
      a.href = link(a.getAttribute('data-page'));
    });
  }
  document.addEventListener('DOMContentLoaded', setupLinks);

  return {
    API_URL: API_URL,
    isConfigured: isConfigured,
    link: link,
    absoluteLink: absoluteLink,
    get: get,
    post: post,
    esc: esc,
    session: session,
    local: local,
    toast: toast,
    pinGate: pinGate,
    logout: logout
  };
})();
