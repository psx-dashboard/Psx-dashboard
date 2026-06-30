
  import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
  import {
    getAuth,
    onAuthStateChanged,
    signInWithEmailAndPassword,
    createUserWithEmailAndPassword,
    signOut,
    GoogleAuthProvider,
    signInWithPopup
  } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-auth.js";
  import {
    getFirestore,
    doc,
    getDoc,
    setDoc
  } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore.js";

  // ===== ADMIN ALLOWLIST =====
  // Only signed-in users whose email is in this list see the Data menu
  // (Update Data / Save Data JSON / Save Full HTML). Everyone else who signs
  // in sees the dashboard normally, just without that menu.
  // >>> REPLACE the email below with your real sign-in email before going live. <<<
  const ADMIN_EMAILS = ['bilal1947@gmail.com'];

  // Your web app's Firebase configuration
  const firebaseConfig = {
    apiKey: "AIzaSyCmIUoD99B2U94wWuHIaQIWmU2A4kppbDY",
    authDomain: "psx-dashboard-dev.firebaseapp.com",
    projectId: "psx-dashboard-dev",
    storageBucket: "psx-dashboard-dev.firebasestorage.app",
    messagingSenderId: "1089260456151",
    appId: "1:1089260456151:web:38c68733e2f4d547330892"
  };

  const app = initializeApp(firebaseConfig);
  const db = getFirestore(app);

  // ===== Watchlist sync bridge (Firestore) =====
  // The watchlist UI/state lives in a separate, non-module <script> earlier in
  // this file (it predates Firebase Auth and doesn't import any SDK). Rather
  // than duplicate Firestore imports there, we expose two small async
  // functions on window that it can call directly. Each user's watchlist is
  // stored at users/{uid}/data/watchlist — see the matching security rule
  // note further down.
  window.fsLoadWatchlist = async function (uid) {
    try {
      const snap = await getDoc(doc(db, 'users', uid, 'data', 'watchlist'));
      return snap.exists() ? (snap.data().tickers || []) : null; // null = no doc yet
    } catch (e) {
      console.error('Firestore load watchlist failed:', e);
      return undefined; // undefined = read failed (distinct from "no doc yet")
    }
  };
  window.fsSaveWatchlist = async function (uid, tickers) {
    try {
      await setDoc(doc(db, 'users', uid, 'data', 'watchlist'), { tickers, updatedAt: Date.now() });
      return true;
    } catch (e) {
      console.error('Firestore save watchlist failed:', e);
      return false;
    }
  };

  // ===== Auto-refresh on return =====
  // Data is baked into this HTML at publish time, so the only way to pick up
  // a newer version is a full page reload. If the user leaves this tab open
  // (or backgrounds the mobile browser) and comes back after a while, this
  // silently reloads so they land on whatever's currently published instead
  // of stale in-memory data.
  const PAGE_LOAD_TIME = Date.now();
  const AUTO_REFRESH_INTERVAL_MS = 60 * 60 * 1000; // 1 hour
  function refreshIfStale() {
    if (document.visibilityState === 'visible' && Date.now() - PAGE_LOAD_TIME > AUTO_REFRESH_INTERVAL_MS) {
      window.location.reload();
    }
  }
  document.addEventListener('visibilitychange', refreshIfStale);
  window.addEventListener('focus', refreshIfStale);
  const auth = getAuth(app);

  const overlay = document.getElementById('authOverlay');
  const userMenu = document.getElementById('authUserMenu');
  const avatarBtn = document.getElementById('authAvatarBtn');
  const dropdown = document.getElementById('authDropdown');
  const avatarLg = document.getElementById('authAvatarLg');
  const dropdownName = document.getElementById('authDropdownName');
  const dropdownEmail = document.getElementById('authDropdownEmail');
  const alertsBtn = document.getElementById('authAlertsBtn');
  const alertsPanel = document.getElementById('alertsPanel');
  const signOutItem = document.getElementById('authSignOutItem');
  const emailEl = document.getElementById('authEmail');
  const passEl = document.getElementById('authPassword');
  const errEl = document.getElementById('authError');
  const submitBtn = document.getElementById('authSubmitBtn');
  const toggleBtn = document.getElementById('authToggleMode');
  const titleEl = document.getElementById('authTitle');
  const subEl = document.getElementById('authSub');

  let mode = 'signin'; // 'signin' | 'signup'

  function setMode(next) {
    mode = next;
    errEl.textContent = '';
    if (mode === 'signin') {
      titleEl.textContent = 'Sign in';
      subEl.textContent = 'Sign in to access Nexus PSX';
      submitBtn.textContent = 'Sign in';
      toggleBtn.textContent = "Need an account? Sign up";
    } else {
      titleEl.textContent = 'Create account';
      subEl.textContent = 'Sign up to access the PSX dashboard';
      submitBtn.textContent = 'Sign up';
      toggleBtn.textContent = 'Already have an account? Sign in';
    }
  }

  toggleBtn.addEventListener('click', () => setMode(mode === 'signin' ? 'signup' : 'signin'));

  function friendlyError(err) {
    const code = err && err.code || '';
    if (code.includes('invalid-email')) return 'That email address looks invalid.';
    if (code.includes('user-not-found') || code.includes('wrong-password') || code.includes('invalid-credential')) return 'Incorrect email or password.';
    if (code.includes('email-already-in-use')) return 'An account already exists for that email.';
    if (code.includes('weak-password')) return 'Password should be at least 6 characters.';
    if (code.includes('unauthorized-domain')) return 'This domain is not yet authorized in Firebase (Authentication → Settings → Authorized domains). Add this domain there.';
    if (code.includes('popup-blocked')) return 'Your browser blocked the sign-in popup. Allow popups for this site and try again.';
    if (code.includes('operation-not-allowed')) return 'Google sign-in is not enabled for this project in Firebase (Authentication → Sign-in method).';
    if (code.includes('network-request-failed')) return 'Network error — check your connection and try again.';
    return 'Something went wrong (' + (code || 'unknown error') + '). Please try again.';
  }

  submitBtn.addEventListener('click', async () => {
    errEl.textContent = '';
    const email = emailEl.value.trim();
    const password = passEl.value;
    if (!email || !password) { errEl.textContent = 'Please enter an email and password.'; return; }
    submitBtn.disabled = true;
    try {
      if (mode === 'signin') {
        await signInWithEmailAndPassword(auth, email, password);
      } else {
        await createUserWithEmailAndPassword(auth, email, password);
      }
    } catch (err) {
      errEl.textContent = friendlyError(err);
    } finally {
      submitBtn.disabled = false;
    }
  });

  // Allow pressing Enter in either field to submit
  [emailEl, passEl].forEach(el => el.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') submitBtn.click();
  }));

  avatarBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    dropdown.classList.toggle('hidden');
  });
  document.addEventListener('click', (e) => {
    if (!userMenu.contains(e.target)) dropdown.classList.add('hidden');
  });
  signOutItem.addEventListener('click', () => { dropdown.classList.add('hidden'); signOut(auth); });
  alertsBtn.addEventListener('click', () => {
    dropdown.classList.add('hidden');
    alertsPanel.classList.toggle('hidden');
  });
  document.addEventListener('click', (e) => {
    if (alertsPanel.contains(e.target) || alertsBtn.contains(e.target)) return;
    alertsPanel.classList.add('hidden');
  });
  document.getElementById('alertsPanelClose').addEventListener('click', () => alertsPanel.classList.add('hidden'));

  // ===== Buy-signal alerts: fires when a stock's status CHANGES into the
  // "buy" family (Initial buy signal / Fresh buy signal / Continuation buy
  // signal / their cautious variants — codes 1.5, 2, 2.5 in SIGNAL_STATUS_MAP),
  // not based on Signal date (which stays as historical data and is never
  // used for alert matching). Each ticker's last-seen status code is kept in
  // localStorage so a stock already sitting in a buy signal doesn't keep
  // re-alerting every time the page loads — only a genuine change does.
  const ALERTS_WL_ONLY_KEY = 'psx_alerts_watchlist_only';
  const ALERT_STATUS_BASELINE_KEY = 'psx_alert_status_baseline_v1';
  const BUY_SIGNAL_CODES = new Set([1.5, 2, 2.5]); // Initial / Fresh / Continuation / cautious variants
  const watchlistOnlyToggle = document.getElementById('alertsWatchlistOnlyToggle');
  if (watchlistOnlyToggle) {
    watchlistOnlyToggle.checked = localStorage.getItem(ALERTS_WL_ONLY_KEY) === '1';
    watchlistOnlyToggle.addEventListener('change', () => {
      try { localStorage.setItem(ALERTS_WL_ONLY_KEY, watchlistOnlyToggle.checked ? '1' : '0'); } catch {}
      checkFreshSignalsToday();
    });
  }

  function readStatusBaseline() {
    try { return JSON.parse(localStorage.getItem(ALERT_STATUS_BASELINE_KEY) || '{}'); }
    catch { return {}; }
  }
  function writeStatusBaseline(map) {
    try { localStorage.setItem(ALERT_STATUS_BASELINE_KEY, JSON.stringify(map)); } catch {}
  }

  function findFreshSignalsToday() {
    // Read via window.SOURCE_DATA, not the bare identifier — SOURCE_DATA is
    // declared with const in app.js (a classic script), and module scripts
    // cannot see another script's top-level const/let bindings directly,
    // only true window properties (which app.js explicitly exposes this as).
    if (typeof window.SOURCE_DATA === 'undefined' || !Array.isArray(window.SOURCE_DATA) || !window.SOURCE_DATA.length) return [];
    if (typeof sigStatusCode !== 'function') return []; // signal-status helpers not loaded yet

    const baseline = readStatusBaseline();
    const isFirstRunEver = Object.keys(baseline).length === 0;
    const newBaseline = {};
    let signals = [];

    window.SOURCE_DATA.forEach(d => {
      const ticker = String(d.Ticker || '');
      if (!ticker) return;
      const code = sigStatusCode(d['Signal Status']);
      newBaseline[ticker] = code;
      if (code == null || !BUY_SIGNAL_CODES.has(code)) return;
      const prevCode = baseline[ticker];
      const wasAlreadyBuySignal = prevCode != null && BUY_SIGNAL_CODES.has(prevCode);
      // First time this browser has ever checked: establish the baseline
      // silently rather than alerting on every stock already sitting in a
      // buy signal — only genuine changes from here on should alert.
      if (!isFirstRunEver && !wasAlreadyBuySignal) signals.push(d);
    });

    writeStatusBaseline(newBaseline);

    const watchlistOnly = watchlistOnlyToggle?.checked;
    if (watchlistOnly && typeof window.getWatchlistTickers === 'function') {
      const tickers = new Set(window.getWatchlistTickers());
      signals = signals.filter(d => tickers.has(String(d.Ticker)));
    }
    return signals;
  }

  function renderAlertsPanel(signals) {
    const body = document.getElementById('alertsPanelBody');
    if (!body) { console.warn('renderAlertsPanel: alertsPanelBody not found in DOM'); return; }
    if (!signals.length) {
      body.innerHTML = '<div class="alerts-panel-empty">No new buy-signal changes since you last checked. Check back after the next data update.</div>';
      return;
    }
    body.innerHTML = signals.map(d => `
      <div class="alert-row">
        <div class="alert-row-left">
          <span class="alert-row-ticker">${d.Ticker}</span>
          <span class="alert-row-name">${d.Name || ''}</span>
        </div>
        <span class="alert-row-tag">${(typeof sigStatusLabel === 'function' ? sigStatusLabel(d['Signal Status']) : null) || d['Signal Status']}</span>
      </div>
    `).join('');
  }

  function showAlertToast(signals) {
    const existing = document.getElementById('alertToast');
    if (existing) existing.remove();
    const toast = document.createElement('div');
    toast.id = 'alertToast';
    toast.className = 'alert-toast';
    const tickers = signals.slice(0, 4).map(d => d.Ticker).join(', ');
    const more = signals.length > 4 ? ` +${signals.length - 4} more` : '';
    toast.innerHTML = `
      <button class="alert-toast-close">✕</button>
      <div class="alert-toast-title" style="display:flex;align-items:center;gap:6px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;"><path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"></path><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"></path></svg>${signals.length} new buy signal${signals.length > 1 ? 's' : ''}</div>
      <div class="alert-toast-body">${tickers}${more}</div>
    `;
    document.body.appendChild(toast);
    toast.querySelector('.alert-toast-close').addEventListener('click', () => toast.remove());
    setTimeout(() => { if (toast.parentNode) toast.remove(); }, 12000);
  }

  function maybeShowBrowserNotification(signals) {
    // Best-effort desktop/mobile popup while the tab is open. True background
    // push (delivered even when the app/tab is closed) needs a service worker
    // + push subscription + a backend (e.g. Firebase Cloud Messaging) — a
    // separate, bigger build than this in-page notification.
    if (!('Notification' in window)) return;
    const fire = () => {
      const tickers = signals.slice(0, 4).map(d => d.Ticker).join(', ');
      new Notification('New buy signal' + (signals.length > 1 ? 's' : ''), {
        body: `${signals.length} stock${signals.length > 1 ? 's' : ''} just turned bullish: ${tickers}`,
        icon: undefined
      });
    };
    if (Notification.permission === 'granted') {
      fire();
    } else if (Notification.permission !== 'denied') {
      Notification.requestPermission().then(perm => { if (perm === 'granted') fire(); });
    }
  }

  function checkFreshSignalsToday() {
    const signals = findFreshSignalsToday();
    const badge = document.getElementById('authAlertBadge');
    const countEl = document.getElementById('authAlertsCount');
    // Defensive: these elements should always exist, but if anything ever
    // removes/renames them, fail quietly rather than throwing — an error
    // here must never be able to block unrelated code (e.g. watchlist sync)
    // that happens to run right after this in the same call stack.
    if (!badge || !countEl) {
      console.warn('checkFreshSignalsToday: authAlertBadge/authAlertsCount not found in DOM');
      renderAlertsPanel(signals);
      return;
    }
    if (signals.length > 0) {
      badge.textContent = signals.length;
      badge.classList.remove('hidden');
      countEl.textContent = signals.length;
      countEl.classList.remove('hidden');
      showAlertToast(signals);
      maybeShowBrowserNotification(signals);
    } else {
      badge.classList.add('hidden');
      countEl.classList.add('hidden');
    }
    renderAlertsPanel(signals);
  }
  window.checkFreshSignalsToday = checkFreshSignalsToday;


  const googleBtn = document.getElementById('authGoogleBtn');
  const googleProvider = new GoogleAuthProvider();
  googleBtn.addEventListener('click', async () => {
    errEl.textContent = '';
    googleBtn.disabled = true;
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (err) {
      if (err && err.code === 'auth/popup-closed-by-user') {
        // user just closed the popup, no need to show an error
      } else {
        errEl.textContent = friendlyError(err);
      }
    } finally {
      googleBtn.disabled = false;
    }
  });

  onAuthStateChanged(auth, (user) => {
    if (user) {
      document.body.classList.remove('authgate-locked');
      overlay.classList.add('hidden');
      const displayName = user.displayName || '';
      const email = user.email || '';
      const initial = (displayName || email || '?').charAt(0).toUpperCase();
      avatarBtn.textContent = initial;
      avatarLg.textContent = initial;
      dropdownName.textContent = displayName || email || 'Account';
      dropdownEmail.textContent = email;
      userMenu.style.display = 'block';
      const isAdmin = ADMIN_EMAILS.includes(email);
      const adminControlsEl = document.getElementById('adminControls');
      if (adminControlsEl) adminControlsEl.style.display = isAdmin ? 'flex' : 'none';
      // Watchlist sync runs first and is wrapped defensively: it must not be
      // skipped just because an unrelated Alerts error happens to throw on
      // the line before it (this previously caused the watchlist to never
      // sync to Firestore at all whenever checkFreshSignalsToday() threw).
      if (typeof window.wlOnSignIn === 'function') window.wlOnSignIn(user.uid, email);
      try { checkFreshSignalsToday(); } catch (e) { console.error('checkFreshSignalsToday failed:', e); }
    } else {
      document.body.classList.add('authgate-locked');
      overlay.classList.remove('hidden');
      userMenu.style.display = 'none';
      const adminControlsEl2 = document.getElementById('adminControls');
      if (adminControlsEl2) adminControlsEl2.style.display = 'none';
      dropdown.classList.add('hidden');
      alertsPanel.classList.add('hidden');
      const toast = document.getElementById('alertToast');
      if (toast) toast.remove();
      if (typeof window.wlOnSignOut === 'function') window.wlOnSignOut();
    }
  });
