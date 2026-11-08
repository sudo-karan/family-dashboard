/* Family FD Tracker — the app's screens. Their source is src/app.jsx (JSX);
 * index.html carries a compiled copy (plain JS, between its GENERATED APP
 * markers), so phones never download Babel or compile anything at startup.
 * To change the app: edit src/app.jsx, run `npm run build`, commit both files.
 * The helpers used here (FDMath, FDPrefs, demo data, APPS_SCRIPT_URL…) live in
 * index.html's first script, which is still edited in place. */
const { useState, useMemo, useEffect, useRef } = React;
const M = FDMath;
const LIVE = !!APPS_SCRIPT_URL;

/* POST to Apps Script. text/plain keeps it a "simple" CORS request (no
 * preflight, which Apps Script cannot answer). */
async function postApi(body){
  let res;
  try {
    res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    });
  } catch (e) {
    throw new Error('network error — are you online?');
  }
  let json;
  try { json = await res.json(); }
  catch (e) { throw new Error('unexpected reply from Apps Script (HTTP ' + res.status + ')'); }
  if (!json.ok) throw new Error(json.error || 'unknown error');
  return json;
}

function daysFromToday(iso){
  const d = M.diffDays(M.todayISO(), iso);
  return d === null ? null : d;
}

/* ---- Web Push (maturity reminders to the installed PWA) ---- */
const PUSH_SUPPORTED = typeof navigator !== 'undefined' && 'serviceWorker' in navigator &&
  typeof window !== 'undefined' && 'PushManager' in window && 'Notification' in window;

function urlB64ToUint8Array(base64){
  const pad = '='.repeat((4 - base64.length % 4) % 4);
  const b64 = (base64 + pad).replace(/-/g, '+').replace(/_/g, '/');
  const raw = atob(b64);
  const out = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i);
  return out;
}

async function getPushSubscription(){
  if (!PUSH_SUPPORTED) return null;
  const reg = await navigator.serviceWorker.ready;
  return reg.pushManager.getSubscription();
}

async function thisDeviceEndpoint(){
  const sub = await getPushSubscription();
  return sub ? sub.endpoint : '';
}

async function enablePush(vapidPublic, device){
  if (!PUSH_SUPPORTED) throw new Error('this device/browser does not support notifications');
  if (!vapidPublic) throw new Error('reminders aren’t set up on the server yet (VAPID keys missing — see SETUP.md)');
  const perm = await Notification.requestPermission();
  if (perm !== 'granted') throw new Error(perm === 'denied'
    ? 'notifications are blocked — allow them for this site in your browser settings'
    : 'notification permission was not granted');
  const reg = await navigator.serviceWorker.ready;
  let sub = await reg.pushManager.getSubscription();
  if (!sub){
    sub = await reg.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: urlB64ToUint8Array(vapidPublic)
    });
  }
  return postApi({ action: 'subscribe', password: pushAuth(), device: device || FDPrefs.getDevice(),
    subscription: sub.toJSON() });
}

/* turn off this device: drop its browser subscription, then remove it server-side */
async function disablePush(){
  const sub = await getPushSubscription();
  if (!sub) return postApi({ action: 'unsubscribe', password: pushAuth(), endpoint: '' });
  const endpoint = sub.endpoint;
  try { await sub.unsubscribe(); } catch (e) { /* best effort */ }
  return postApi({ action: 'unsubscribe', password: pushAuth(), endpoint: endpoint });
}

/* remove any device from the list by its endpoint (used by the Unsubscribe
 * button next to each device on the Reminders tab) */
async function unsubscribeEndpoint(endpoint){
  return postApi({ action: 'unsubscribe', password: pushAuth(), endpoint: endpoint });
}

/* the shared password, captured after login so push helpers can authenticate */
let _pushAuth = '';
function pushAuth(){ return _pushAuth; }
function setPushAuth(pw){ _pushAuth = pw; }

function interestOf(fd){
  return (fd.maturity > 0) ? fd.maturity - fd.principal : null;
}

/* Inactive FDs never count toward the dashboard; the flag applies to Active FDs. */
function isOnDashboard(fd){
  return fd.status === 'Active' && !!fd.showInDashboard;
}

/* Has an Active FD's end date already passed? Such an FD stays Active and
 * reads "Matured Nd ago" until someone taps "Mark as done" — it is never
 * archived automatically. */
function isMatured(fd){
  if (fd.status !== 'Active' || !M.parseISO(fd.endDate)) return false;
  const d = daysFromToday(fd.endDate); /* endDate − today; negative = overdue */
  return d !== null && d < 0;
}

/* Pill text + colour for the status cell: matured Actives read "Matured Nd ago". */
function statusLabel(fd){
  if (fd.status !== 'Active') return { text: fd.status || 'Inactive', cls: 'gray' };
  if (isMatured(fd)) return { text: 'Matured ' + (-daysFromToday(fd.endDate)) + 'd ago', cls: 'red' };
  return { text: 'Active', cls: 'green' };
}

/* "Days left" cell in All FDs: days until an Active FD's end date, turning red
 * once DAYS_LEFT_RED or fewer remain (and on/after the maturity date). Only a
 * running FD has days left, so Inactive FDs — and any FD without a valid end
 * date — read a muted "—". `cls` reuses the End column's classes. */
const DAYS_LEFT_RED = 10;
function daysLeftOf(fd){
  const d = (fd.status === 'Active' && M.parseISO(fd.endDate)) ? daysFromToday(fd.endDate) : null;
  if (d === null) return { text: '—', cls: ' muted' };
  if (d < 0) return { text: 'Matured', cls: ' date-red' };
  if (d === 0) return { text: 'Today', cls: ' date-red' };
  return { text: d + (d === 1 ? ' day' : ' days'), cls: d <= DAYS_LEFT_RED ? ' date-red' : '' };
}

function tenureCellOf(fd){
  if (M.parseISO(fd.startDate) && M.parseISO(fd.endDate)){
    return { label: M.tenureLabel(M.diffYMD(fd.startDate, fd.endDate)), days: M.diffDays(fd.startDate, fd.endDate) };
  }
  if (fd.tenureDays > 0) return { label: fd.tenureDays + 'd', days: fd.tenureDays };
  return { label: '—', days: null };
}

/* Everything searchable about an FD, in every spelling someone might type:
 * amounts as raw digits, with Indian commas, and in words; dates as stored
 * and as displayed. Fields are joined with '|' so a space-insensitive query
 * can't accidentally match across two fields. */
function searchHay(f){
  const parts = [
    f.account, f.person, f.bank, String(f.accountNumber || ''), f.remarks || '',
    f.status, f.interestOnMaturity ? 'on maturity' : 'periodic payout', f.compounding,
    f.startDate || '', M.fmtDate(f.startDate), f.endDate || '', M.fmtDate(f.endDate),
    f.rate ? f.rate + '% ' + f.rate : '',
    tenureCellOf(f).label, f.tenureDays ? f.tenureDays + ' days' : '',
    '#' + f.id
  ];
  [f.principal, f.maturity, interestOf(f)].forEach(function(v){
    if (v === null || v === undefined || v === '' || v === 0) return;
    const r = M.round2(Number(v));
    parts.push(String(r));            /* 200000 */
    parts.push(M.fmtINR(r));          /* ₹2,00,000 */
    const w = M.fmtINRWords(r);
    if (w) parts.push(w);             /* 2 lakh */
  });
  return parts.join('|').toLowerCase();
}

/* Match a query against a haystack, ignoring ₹, commas and spaces — so
 * "200000", "2,00,000", "₹2,00,000" and "2 lakh" all find the same FD. */
function searchMatch(hay, q){
  const needle = String(q || '').trim().toLowerCase();
  if (!needle) return true;
  if (hay.indexOf(needle) >= 0) return true;
  const strip = function(s){ return s.replace(/[₹,\s]/g, ''); };
  const n = strip(needle);
  return n !== '' && strip(hay).indexOf(n) >= 0;
}

/* ============================== login ============================== */
function Login({ onLogin, busy, error }){
  const [pw, setPw] = useState('');
  const [dev, setDev] = useState(function(){ return FDPrefs.getDevice(); });
  return (
    <div className="login-wrap">
      <form className="login-card" onSubmit={function(e){ e.preventDefault(); onLogin(pw, dev); }}>
        <div className="login-mark" aria-hidden="true">₹</div>
        <h1>Family FD Tracker</h1>
        <p className="login-sub">Fixed deposits across the family, in one ledger.</p>
        {!LIVE && (
          <p className="login-demo">
            Demo mode — enter any password. The data is a sample and resets on refresh.
          </p>
        )}
        <label htmlFor="pw">Family password</label>
        <input id="pw" type="password" autoFocus autoComplete="current-password"
               value={pw} onChange={function(e){ setPw(e.target.value); }} />
        <label htmlFor="dev">Device name <span className="muted" style={{ fontWeight:500 }}>(optional — shows in the change log)</span></label>
        <input id="dev" type="text" placeholder="e.g. Dad's phone" autoComplete="off"
               value={dev} onChange={function(e){ setDev(e.target.value); }} />
        {error && <p className="login-err" role="alert">{error}</p>}
        <button type="submit" className="btn primary" disabled={busy}>
          {busy ? 'Checking…' : 'Open the ledger'}
        </button>
        <p className="login-sub small" style={{ margin:'12px 0 0' }}>
          Stays signed in on this device until you sign out.
        </p>
      </form>
    </div>
  );
}

/* ============================== dashboard ============================== */
function StatCards({ fds, shown, view }){
  const hidden = view === 'hidden';
  const inactiveCount = fds.filter(function(f){ return f.status !== 'Active'; }).length;
  const hiddenCount = fds.filter(function(f){ return f.status === 'Active' && !f.showInDashboard; }).length;
  let viewNote;
  if (hidden){
    /* the hidden view is only active FDs kept off the dashboard on purpose */
    viewNote = shown.length ? 'active, kept off the dashboard' : 'nothing hidden';
  } else {
    const parts = [];
    if (inactiveCount) parts.push(inactiveCount + ' inactive');
    if (hiddenCount) parts.push(hiddenCount + ' hidden');
    viewNote = parts.length ? 'excludes ' + parts.join(' · ') : 'all deposits counted';
  }
  const tp = shown.reduce(function(s, f){ return s + (Number(f.principal) || 0); }, 0);
  const tm = shown.reduce(function(s, f){ return s + (Number(f.maturity) || 0); }, 0);
  const accounts = {};
  shown.forEach(function(f){ accounts[f.account] = true; });
  const nAcc = Object.keys(accounts).length;
  const sub = function(amount, extra){
    const w = M.fmtINRWords(amount);
    const parts = [];
    if (w) parts.push(w);
    if (extra) parts.push(extra);
    return parts.join(' · ') || ' ';
  };
  return (
    <div className="cards">
      <div className="card">
        <div className="card-label">{hidden ? 'Hidden deposits' : 'Deposits in view'}</div>
        <div className="card-value">{shown.length}</div>
        <div className="card-sub">{viewNote}</div>
      </div>
      <div className="card">
        <div className="card-label">Total principal</div>
        <div className="card-value">{M.fmtINR(tp)}</div>
        <div className="card-sub">{sub(tp, 'across ' + nAcc + ' account' + (nAcc === 1 ? '' : 's'))}</div>
      </div>
      <div className="card">
        <div className="card-label">Total at maturity</div>
        <div className="card-value">{M.fmtINR(tm)}</div>
        <div className="card-sub">{sub(tm, tp > 0 ? '+' + ((tm / tp - 1) * 100).toFixed(1) + '% over principal' : '')}</div>
      </div>
      <div className="card">
        <div className="card-label">Interest to earn</div>
        <div className="card-value accent">{M.fmtINR(tm - tp)}</div>
        <div className="card-sub">{sub(tm - tp, 'maturity − principal')}</div>
      </div>
    </div>
  );
}

function soonPill(days){
  if (days < 0) return <span className="pill red soon-pill">matured {-days}d ago</span>;
  if (days === 0) return <span className="pill amber soon-pill">matures today</span>;
  if (days <= 30) return <span className="pill amber soon-pill">in {days}d</span>;
  return <span className="pill green soon-pill">in {days}d</span>;
}

function MaturingSoon({ shown, onEdit, onMarkDone, flash }){
  const rows = shown
    .filter(function(f){ return f.status === 'Active' && M.parseISO(f.endDate); })
    .map(function(f){ return { fd: f, days: daysFromToday(f.endDate) }; })
    .filter(function(x){ return x.days <= 120; })
    .sort(function(a, b){ return a.days - b.days; });
  return (
    <div className="panel">
      <div className="panel-head">
        <div className="panel-title">Maturing soon</div>
        <div className="panel-sub">Active deposits ending within 120 days, including overdue. Matured? Tap “Mark as done” once you’ve renewed or withdrawn it.</div>
      </div>
      {rows.length === 0 && (
        <div className="panel-empty">Nothing matures in the next 120 days. Sit back and let the interest build.</div>
      )}
      {rows.map(function(x){
        const f = x.fd, t = tenureCellOf(f);
        return (
          <div key={f.id} className={'soon-row' + (isFlashFd(flash, f.id) ? ' flash' : '')} role="button" tabIndex={0}
               onClick={function(ev){ if (ev.target === ev.currentTarget || !ev.target.closest('button')) onEdit(f); }}
               onKeyDown={function(ev){ if (ev.target === ev.currentTarget && (ev.key === 'Enter' || ev.key === ' ')){ ev.preventDefault(); onEdit(f); } }}>
            {soonPill(x.days)}
            <span className="soon-main">
              <span className="soon-acct">{f.account}{f.remarks ? <span className="muted small"> · {f.remarks}</span> : null}</span>
              <span className="soon-meta">ends {M.fmtDate(f.endDate)} · {M.fmtRate(f.rate)} · {t.label}</span>
            </span>
            <span className="soon-amts num">
              <span className="soon-mat">{M.fmtINR(f.maturity)}</span>
              {M.fmtINRWords(f.maturity) ? <span className="soon-words">{M.fmtINRWords(f.maturity)}</span> : null}
              <span className="soon-pr">on {M.fmtINR(f.principal)}</span>
            </span>
            {x.days < 0 ? (
              <button type="button" className="btn done-btn"
                      title="Move this matured FD to Inactive"
                      onClick={function(ev){ ev.stopPropagation(); onMarkDone(f); }}>Mark as done</button>
            ) : null}
          </div>
        );
      })}
    </div>
  );
}

function Breakdown({ title, rows }){
  const max = rows.reduce(function(m, r){ return Math.max(m, r.principal); }, 0);
  return (
    <div className="panel">
      <div className="panel-head"><div className="panel-title">{title}</div></div>
      {rows.length === 0 ? <div className="panel-empty">Nothing to show.</div> : (
       <div className="mini-scroll">
        <table className="mini-table">
          <thead>
            <tr><th>{title === 'By person' ? 'Person' : 'Account'}</th><th className="num">FDs</th><th className="num">Principal</th><th className="num">Maturity</th></tr>
          </thead>
          <tbody>
            {rows.map(function(r){
              return (
                <tr key={r.key}>
                  <td>
                    <div className="share-name">{r.key}</div>
                    {r.sub ? <div className="cell-sub">{r.sub}</div> : null}
                    <div className="share-track"><div className="share-bar" style={{ width: (max > 0 ? (r.principal / max) * 100 : 0) + '%' }} /></div>
                  </td>
                  <td className="num">{r.count}</td>
                  <td className="num">{M.fmtINR(r.principal)}{M.fmtINRWords(r.principal) ? <div className="cell-sub cell-words">{M.fmtINRWords(r.principal)}</div> : null}</td>
                  <td className="num">{M.fmtINR(r.maturity)}{M.fmtINRWords(r.maturity) ? <div className="cell-sub cell-words">{M.fmtINRWords(r.maturity)}</div> : null}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
       </div>
      )}
    </div>
  );
}

/* ============================== recent changes ============================== */
function fmtStamp(s){
  const m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}:\d{2})/.exec(String(s || ''));
  if (!m) return String(s || '');
  const iso = m[1] + '-' + m[2] + '-' + m[3];
  if (iso === M.todayISO()) return 'Today ' + m[4];
  return M.fmtDate(iso).replace(/\s\d{4}$/, '') + ', ' + m[4];
}

/* Saved-change highlight helpers (flash = {fds:[ids], account:name|null}). */
function isFlashFd(flash, id){ return !!(flash && flash.fds && flash.fds.indexOf(id) >= 0); }
function isFlashAccount(flash, name){
  return !!(flash && flash.account && String(flash.account).trim().toLowerCase() === String(name || '').trim().toLowerCase());
}
/* What an undo touched, from its log target ('fd:12' / 'account:mom sbi'). */
function flashFromTarget(target){
  const f = /^fd:(\d+)$/.exec(String(target || ''));
  if (f) return { fds: [parseInt(f[1], 10)], account: null };
  const a = /^account:(.+)$/.exec(String(target || ''));
  return a ? { fds: [], account: a[1] } : null;
}

function logPill(action){
  if (action === 'create' || action === 'accountCreate') return <span className="pill green log-pill">added</span>;
  if (action === 'update' || action === 'accountUpdate') return <span className="pill amber log-pill">edited</span>;
  if (action === 'delete' || action === 'accountDelete') return <span className="pill red log-pill">deleted</span>;
  if (action === 'transfer') return <span className="pill amber log-pill">moved</span>;
  if (action === 'markDone') return <span className="pill gray log-pill">archived</span>;
  if (action === 'autoArchive') return <span className="pill gray log-pill">auto-archived</span>;
  return <span className="pill gray log-pill">undid</span>;
}

/* An FD/account snapshot the read-only viewer can render, or null (e.g. for
 * transfers, whose payload is {from,to,ids} and whose summary says it all). */
function logSnapshot(e){
  const snap = e.after || e.before;
  if (!snap) return null;
  if (/^fd:/.test(e.target)) return { snap: snap, isFd: true };
  if (/^account:/.test(e.target) && snap.name) return { snap: snap, isFd: false };
  return null;
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
/* 'YYYY-MM' -> 'November 2026'; 'undated' -> 'Undated'. */
function monthLabel(key){
  if (key === 'undated') return 'Undated';
  const m = /^(\d{4})-(\d{2})$/.exec(String(key || ''));
  return m ? MONTH_NAMES[parseInt(m[2], 10) - 1] + ' ' + m[1] : String(key || '');
}

/* History: one month of changes at a time (this month by default), with no
 * cap on how many. `logMonths` lists every month that has changes; `asked` is
 * a month picked but not yet loaded. logMonths === null means the backend
 * predates month browsing: show the latest-15 list it sends, and how to fix. */
function RecentChanges({ log, logMonth, logMonths, loading, wantMonth, onPickMonth, onUndo, onView }){
  const months = Array.isArray(logMonths) ? logMonths : null;
  /* while the month wanted (or "this month") is on its way, nothing stale is
     shown: the list says what's loading and the controls wait */
  const busy = !!(months && loading);
  const busyLabel = wantMonth ? monthLabel(wantMonth) : 'this month';
  const shown = logMonth;
  const idx = months ? months.findIndex(function(m){ return m.month === shown; }) : -1;
  const older = idx >= 0 && idx < months.length - 1 ? months[idx + 1].month : null;
  const newer = idx > 0 ? months[idx - 1].month : null;
  const countText = function(n){ return n ? n + (n === 1 ? ' change' : ' changes') : 'no changes'; };
  return (
    <div className="panel">
      <div className="panel-head">
        <div className="panel-title">Change history</div>
        <div className="panel-sub">Every add, edit and delete, newest first{months ? ', one month at a time' : ''}. Tap an entry to see exactly what was recorded. The latest change to an item can be undone.</div>
      </div>
      {months ? (
        <div className="filters log-months" role="group" aria-label="Choose a month">
          <button type="button" className="btn" disabled={!older || busy}
                  onClick={function(){ onPickMonth(older); }}>‹ Older</button>
          {busy ? (
            <select aria-label="Month" value="" disabled>
              <option value="">{busyLabel.charAt(0).toUpperCase() + busyLabel.slice(1)}…</option>
            </select>
          ) : (
            <select aria-label="Month" value={shown || ''}
                    onChange={function(e){ onPickMonth(e.target.value); }}>
              {months.map(function(m){
                return <option key={m.month} value={m.month}>{monthLabel(m.month) + ' — ' + countText(m.count)}</option>;
              })}
            </select>
          )}
          <button type="button" className="btn" disabled={!newer || busy}
                  onClick={function(){ onPickMonth(newer); }}>Newer ›</button>
        </div>
      ) : (
        <div className="log-legacy">Showing only the 15 most recent changes. To browse every month, paste the latest <b>Code.gs</b> into Apps Script and deploy a new version.</div>
      )}
      {busy ? (
        <div className="panel-empty" role="status">Loading {busyLabel}…</div>
      ) : (!log || log.length === 0) ? (
        <div className="panel-empty">{months
          ? 'No changes in ' + monthLabel(shown) + '.' + (older ? ' Tap ‹ Older to see earlier months.' : '')
          : 'No changes recorded yet — they\u2019ll show up here as the family edits.'}</div>
      ) : log.map(function(e){
        const viewable = !!logSnapshot(e);
        const inner = (
          <React.Fragment>
            {logPill(e.action)}
            <span className="log-main">
              <span className="log-sum">{e.summary}</span>
              <span className="log-meta">{fmtStamp(e.time)}{e.device ? ' · ' + e.device : ''}{e.undone ? ' · undone' : ''}</span>
            </span>
            {e.canUndo ? (
              <button type="button" className="link-btn"
                      onClick={function(ev){ ev.stopPropagation(); onUndo(e); }}>Undo</button>
            ) : null}
            {viewable ? <span className="log-chev" aria-hidden="true">›</span> : null}
          </React.Fragment>
        );
        return viewable ? (
          <div key={e.id} className="log-row clickable" role="button" tabIndex={0}
               title="See the full details of this change"
               onClick={function(ev){ if (ev.target === ev.currentTarget || !ev.target.closest('button')) onView(e); }}
               onKeyDown={function(ev){ if (ev.target === ev.currentTarget && (ev.key === 'Enter' || ev.key === ' ')){ ev.preventDefault(); onView(e); } }}>
            {inner}
          </div>
        ) : (
          <div key={e.id} className="log-row">{inner}</div>
        );
      })}
    </div>
  );
}

/* ============================== reminders tab ============================== */
function RemindersView({ pushSubs, pushState, thisEndpoint, busy, onEnable, onDisable, onUnsubscribe }){
  const subs = pushSubs || [];
  const statusLine = {
    on: 'Reminders are ON for this device.',
    off: 'Reminders are off on this device.',
    denied: 'Notifications are blocked for this site — allow them in your browser settings, then reload.',
    unsupported: 'This device or browser can’t show push notifications (on iPhone, add the app to your Home Screen first, iOS 16.4+).',
    busy: 'Working…'
  }[pushState] || '';
  return (
    <div>
      <div className="panel">
        <div className="panel-head">
          <div className="panel-title">Maturity reminders</div>
          <div className="panel-sub">A phone notification 2 days before, 1 day before, and on the maturity date (around 10 AM IST). Turn it on once per device.</div>
        </div>
        <div className="reminders-status">
          <span className={'rem-dot ' + (pushState === 'on' ? 'on' : pushState === 'denied' ? 'bad' : 'off')} aria-hidden="true" />
          <span className="rem-status-text">{statusLine}</span>
          {pushState === 'on' ? (
            <button type="button" className="btn" disabled={busy} onClick={onDisable}>Turn off on this device</button>
          ) : pushState === 'off' ? (
            <button type="button" className="btn primary" disabled={busy} onClick={onEnable}>Enable on this device</button>
          ) : pushState === 'busy' ? (
            <button type="button" className="btn" disabled={true}>Working…</button>
          ) : null}
        </div>
      </div>

      <div className="panel">
        <div className="panel-head">
          <div className="panel-title">Subscribed devices</div>
          <div className="panel-sub">Every device that will receive reminders. Names come from the “Device name” entered at sign-in.</div>
        </div>
        {subs.length === 0 ? (
          <div className="panel-empty">No devices subscribed yet — tap “Enable on this device” above.</div>
        ) : (
         <div className="mini-scroll">
          <table className="mini-table">
            <thead><tr><th>Device</th><th>Subscribed</th><th aria-label="Actions"></th></tr></thead>
            <tbody>
              {subs.map(function(s){
                const isThis = thisEndpoint && s.endpoint === thisEndpoint;
                return (
                  <tr key={s.endpoint}>
                    <td>
                      <span className="share-name">{s.device ? s.device : 'Unnamed device'}</span>
                      {isThis ? <span className="pill green" style={{ marginLeft:8 }}>this device</span> : null}
                      <div className="cell-sub" style={{ whiteSpace:'normal', wordBreak:'break-all', maxWidth:260 }}>…{String(s.endpoint).slice(-22)}</div>
                    </td>
                    <td className="date-cell">{M.fmtDate(s.added)}</td>
                    <td className="row-actions" style={{ textAlign:'right' }}>
                      <button type="button" className="btn danger-ghost small" disabled={busy}
                              onClick={function(){ onUnsubscribe(s); }}>Unsubscribe</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
         </div>
        )}
      </div>
    </div>
  );
}

function Dashboard({ fds, onEdit, onMarkDone, flash }){
  const [view, setView] = useState('counted'); /* 'counted' | 'hidden' */
  const counted = fds.filter(isOnDashboard);
  /* "hidden" = Active deposits deliberately kept off the dashboard
     (showInDashboard=false). Inactive ones don't get a summary at all. */
  const hidden = fds.filter(function(f){ return f.status === 'Active' && !f.showInDashboard; });
  const shown = view === 'hidden' ? hidden : counted;
  function group(keyFn, subFn){
    const map = {};
    shown.forEach(function(f){
      const k = keyFn(f) || '—';
      if (!map[k]) map[k] = { key:k, sub: subFn ? subFn(f) : '', count:0, principal:0, maturity:0 };
      map[k].count += 1;
      map[k].principal += Number(f.principal) || 0;
      map[k].maturity += Number(f.maturity) || 0;
    });
    return Object.keys(map).map(function(k){ return map[k]; })
      .sort(function(a, b){ return b.principal - a.principal; });
  }
  return (
    <div>
      <div className="dash-toggle">
        <div className="seg" role="group" aria-label="Which deposits to summarise">
          <button type="button" className={'seg-btn' + (view === 'counted' ? ' on' : '')}
                  onClick={function(){ setView('counted'); }}>On dashboard ({counted.length})</button>
          <button type="button" className={'seg-btn' + (view === 'hidden' ? ' on' : '')}
                  onClick={function(){ setView('hidden'); }}>Hidden ({hidden.length})</button>
        </div>
      </div>
      <StatCards fds={fds} shown={shown} view={view} />
      <MaturingSoon shown={shown} onEdit={onEdit} onMarkDone={onMarkDone} flash={flash} />
      <div className="grid-2">
        <Breakdown title="By person" rows={group(function(f){ return f.person; })} />
        <Breakdown title="By account" rows={group(function(f){ return f.account; }, function(f){ return f.person; })} />
      </div>
    </div>
  );
}

/* ============================== all FDs table ============================== */
function exportCSV(rows){
  const head = ['ID','Account','Person','Bank','FD Account Number','Start Date','End Date',
    'Tenure (Days)','Interest %','Principal','Maturity','Interest','Payout','Compounding',
    'Status','Show In Dashboard','Remarks'];
  const esc = function(v){
    const s = String(v === null || v === undefined ? '' : v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [head.map(esc).join(',')];
  rows.forEach(function(f){
    const i = interestOf(f);
    lines.push([f.id, f.account, f.person, f.bank, f.accountNumber, f.startDate, f.endDate,
      f.tenureDays || '', f.rate || '', f.principal || '', f.maturity || '',
      i === null ? '' : M.round2(i), f.interestOnMaturity ? 'On maturity' : 'Periodic',
      f.compounding, f.status, f.showInDashboard ? 'TRUE' : 'FALSE', f.remarks].map(esc).join(','));
  });
  const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'fd-export-' + M.todayISO() + '.csv';
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

const SORTS = {
  'end-asc':       { label: 'Maturity: soonest first' },
  'end-desc':      { label: 'Maturity: latest first' },
  'start-desc':    { label: 'Start date: newest first' },
  'principal-desc':{ label: 'Principal: high → low' },
  'rate-desc':     { label: 'Rate: high → low' }
};

function AllFDs({ fds, accounts, onEdit, onDelete, onMarkDone, flash }){
  const [person, setPerson] = useState('all');
  const [account, setAccount] = useState('all');
  const [status, setStatus] = useState('Active'); /* default: show live FDs */
  const [dash, setDash] = useState('all');
  const [q, setQ] = useState('');
  const [sort, setSort] = useState('end-asc');
  /* after a save, bring the highlighted row into view */
  const flashKey = flash && flash.fds && flash.fds.length ? flash.n + ':' + flash.fds.join(',') : '';
  useEffect(function(){
    if (!flashKey) return;
    const row = document.querySelector('.fd-table tr.flash');
    if (row && row.scrollIntoView) row.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }, [flashKey]);

  const persons = useMemo(function(){
    const seen = {};
    accounts.forEach(function(a){ if (a.person) seen[a.person] = true; });
    return Object.keys(seen).sort();
  }, [accounts]);

  const accountOptions = useMemo(function(){
    return accounts
      .filter(function(a){ return person === 'all' || a.person === person; })
      .map(function(a){ return a.name; })
      .sort();
  }, [accounts, person]);

  function pickPerson(p){
    setPerson(p);
    if (account !== 'all'){
      const acc = accounts.find(function(a){ return a.name === account; });
      if (p !== 'all' && (!acc || acc.person !== p)) setAccount('all');
    }
  }

  const hays = useMemo(function(){
    const map = {};
    fds.forEach(function(f){ map[f.id] = searchHay(f); });
    return map;
  }, [fds]);

  const filtered = useMemo(function(){
    let rows = fds.filter(function(f){
      if (person !== 'all' && f.person !== person) return false;
      if (account !== 'all' && f.account !== account) return false;
      if (status === 'Matured'){ if (!isMatured(f)) return false; }
      else if (status !== 'all' && f.status !== status) return false;
      if (dash === 'shown' && !isOnDashboard(f)) return false;
      if (dash === 'hidden' && isOnDashboard(f)) return false;
      if (!searchMatch(hays[f.id], q)) return false;
      return true;
    });
    const endKey = function(f){ const p = M.parseISO(f.endDate); return p ? Date.UTC(p.y, p.mo - 1, p.d) : null; };
    const startKey = function(f){ const p = M.parseISO(f.startDate); return p ? Date.UTC(p.y, p.mo - 1, p.d) : null; };
    const cmp = {
      'end-asc': function(a, b){
        const ka = endKey(a), kb = endKey(b);
        return (ka === null ? Infinity : ka) - (kb === null ? Infinity : kb);
      },
      'end-desc': function(a, b){
        const ka = endKey(a), kb = endKey(b);
        return (kb === null ? -Infinity : kb) - (ka === null ? -Infinity : ka);
      },
      'start-desc': function(a, b){
        const ka = startKey(a), kb = startKey(b);
        return (kb === null ? -Infinity : kb) - (ka === null ? -Infinity : ka);
      },
      'principal-desc': function(a, b){ return (b.principal || 0) - (a.principal || 0); },
      'rate-desc': function(a, b){ return (b.rate || 0) - (a.rate || 0); }
    }[sort];
    return rows.slice().sort(cmp);
  }, [fds, hays, person, account, status, dash, q, sort]);

  const totP = filtered.reduce(function(s, f){ return s + (Number(f.principal) || 0); }, 0);
  const totM = filtered.reduce(function(s, f){ return s + (Number(f.maturity) || 0); }, 0);

  function endCls(f){
    if (!M.parseISO(f.endDate)) return '';
    const d = daysFromToday(f.endDate);
    if (d < 0) return ' date-red';
    if (d <= 30) return ' date-amber';
    return '';
  }

  return (
    <div className="panel">
      <div className="filters">
        <select aria-label="Filter by person" value={person} onChange={function(e){ pickPerson(e.target.value); }}>
          <option value="all">Everyone</option>
          {persons.map(function(p){ return <option key={p} value={p}>{p}</option>; })}
        </select>
        <select aria-label="Filter by account" value={account} onChange={function(e){ setAccount(e.target.value); }}>
          <option value="all">All accounts</option>
          {accountOptions.map(function(n){ return <option key={n} value={n}>{n}</option>; })}
        </select>
        <select aria-label="Filter by status" value={status} onChange={function(e){ setStatus(e.target.value); }}>
          <option value="all">Any status</option>
          <option value="Active">Active</option>
          <option value="Matured">Matured</option>
          <option value="Inactive">Inactive</option>
        </select>
        <select aria-label="Filter by dashboard visibility" value={dash} onChange={function(e){ setDash(e.target.value); }}>
          <option value="all">Dashboard: all</option>
          <option value="shown">Shown</option>
          <option value="hidden">Hidden</option>
        </select>
        <input type="search" placeholder="Search anything — name, FD no., amount, date…" aria-label="Search"
               value={q} onChange={function(e){ setQ(e.target.value); }} />
        <span className="spacer" />
        <select aria-label="Sort" value={sort} onChange={function(e){ setSort(e.target.value); }}>
          {Object.keys(SORTS).map(function(k){ return <option key={k} value={k}>{SORTS[k].label}</option>; })}
        </select>
        <button type="button" className="btn" onClick={function(){ exportCSV(filtered); }}>Export CSV</button>
      </div>
      <div className="table-scroll">
        <table className="fd-table">
          <thead>
            <tr>
              <th className="col-actions" aria-label="Actions"></th><th className="col-acct">Account</th>
              <th>FD No.</th><th>Start</th><th>End</th><th>Days left</th><th>Tenure</th>
              <th className="num">Rate</th><th className="num">Principal</th><th className="num">Maturity</th>
              <th className="num">Interest</th><th>Payout</th><th>Status</th><th>Dashboard</th>
              <th>Remarks</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={15} className="panel-empty">No deposits match these filters.</td></tr>
            )}
            {filtered.map(function(f){
              const t = tenureCellOf(f);
              const i = interestOf(f);
              const sl = statusLabel(f);
              const dl = daysLeftOf(f);
              return (
                <tr key={f.id} data-fd-id={f.id} className={isFlashFd(flash, f.id) ? 'flash' : undefined}>
                  <td className="col-actions row-actions stacked">
                    <button type="button" className="btn ghost small" onClick={function(){ onEdit(f); }}>Edit</button>
                    {isMatured(f) ? (
                      <button type="button" className="btn ghost small" title="Move this matured FD to Inactive"
                              onClick={function(){ onMarkDone(f); }}>Mark as done</button>
                    ) : null}
                    <button type="button" className="btn danger-ghost small" aria-label={'Delete FD ' + f.id}
                            onClick={function(){ onDelete(f); }}>Delete</button>
                  </td>
                  <td className="col-acct"><div className="cell-main">{f.account}</div><div className="cell-sub">{f.person}</div></td>
                  <td><span className="fdno">{f.accountNumber || '—'}</span></td>
                  <td className="date-cell">{M.fmtDate(f.startDate)}</td>
                  <td className={'date-cell' + endCls(f)}>{M.fmtDate(f.endDate)}</td>
                  <td className={'days-cell' + dl.cls}>{dl.text}</td>
                  <td><div className="cell-main">{t.label}</div><div className="cell-sub num" style={{ textAlign:'left' }}>{t.days !== null ? t.days + ' days' : ''}</div></td>
                  <td className="num">{M.fmtRate(f.rate)}</td>
                  <td className="num">{M.fmtINR(f.principal || '')}{M.fmtINRWords(f.principal) ? <div className="cell-sub cell-words">{M.fmtINRWords(f.principal)}</div> : null}</td>
                  <td className="num">{M.fmtINR(f.maturity || '')}{M.fmtINRWords(f.maturity) ? <div className="cell-sub cell-words">{M.fmtINRWords(f.maturity)}</div> : null}</td>
                  <td className="num">{i === null ? '—' : M.fmtINR(M.round2(i))}{i !== null && M.fmtINRWords(i) ? <div className="cell-sub cell-words">{M.fmtINRWords(M.round2(i))}</div> : null}</td>
                  <td><span className="small muted">{f.interestOnMaturity ? 'On maturity' : 'Periodic'}</span></td>
                  <td><span className={'pill ' + sl.cls}>{sl.text}</span></td>
                  <td><span className={'pill ' + (isOnDashboard(f) ? 'green' : 'gray')}>{isOnDashboard(f) ? 'Shown' : 'Hidden'}</span></td>
                  <td className="remarks-cell" title={f.remarks || ''}>{f.remarks || ''}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            {/* first two cells mirror the pinned body columns, so the count
                stays under Account and each total stays under its real column
                (never sliding left to read like an account total) */}
            <tr>
              <td className="col-actions"></td>
              <td className="col-acct">{filtered.length} of {fds.length} deposits</td>
              <td colSpan={6}></td>
              <td className="num">{M.fmtINR(totP)}{M.fmtINRWords(totP) ? <div className="cell-sub cell-words">{M.fmtINRWords(totP)}</div> : null}</td>
              <td className="num">{M.fmtINR(totM)}{M.fmtINRWords(totM) ? <div className="cell-sub cell-words">{M.fmtINRWords(totM)}</div> : null}</td>
              <td className="num">{M.fmtINR(totM - totP)}{M.fmtINRWords(totM - totP) ? <div className="cell-sub cell-words">{M.fmtINRWords(totM - totP)}</div> : null}</td>
              <td colSpan={4}></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

/* ============================== accounts tab ============================== */
function AccountsView({ accounts, fds, onAdd, onEdit, onDelete, onTransfer, flash }){
  /* Per account: `total` counts ALL FDs (even inactive) for the delete guard,
     and is never shown. Among ACTIVE FDs only, group by remark — an account
     name together with a remark is treated as the grouping key, so FDs sharing
     a remark total together and separately from the untagged ones. */
  const info = {};
  fds.forEach(function(f){
    const a = info[f.account] = info[f.account] || { total:0, groups:{} };
    a.total += 1;
    if (f.status === 'Active'){
      const rk = (f.remarks || '').trim();
      const g = a.groups[rk] = a.groups[rk] || { remark: rk, count:0, principal:0 };
      g.count += 1; g.principal += Number(f.principal) || 0;
    }
  });
  /* remark groups for an account, named ones A→Z with the untagged group last */
  function groupsOf(name){
    const g = (info[name] || {}).groups || {};
    return Object.keys(g).sort(function(x, y){
      return ((!x) - (!y)) || x.localeCompare(y);
    }).map(function(k){ return g[k]; });
  }
  const rows = accounts.slice().sort(function(a, b){
    return (a.person || '').localeCompare(b.person || '') || a.name.localeCompare(b.name);
  });
  return (
    <div className="panel">
      <div className="panel-head" style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:12, flexWrap:'wrap' }}>
        <div>
          <div className="panel-title">Accounts</div>
          <div className="panel-sub">Each account ties its FDs to a person and a bank — what used to be a workbook tab.</div>
        </div>
        <button type="button" className="btn primary" onClick={onAdd}>＋ Add account</button>
      </div>
      {rows.length === 0 ? (
        <div className="panel-empty">No accounts yet — add the first one.</div>
      ) : (
       <div className="mini-scroll">
        <table className="mini-table wide acct-table">
          <thead>
            <tr><th className="col-actions" aria-label="Actions"></th><th className="col-acct">Account</th><th>Person</th><th>Bank</th><th>Remark</th><th className="num">FDs</th><th className="num">Principal</th></tr>
          </thead>
          <tbody>
            {rows.map(function(a){
              const total = (info[a.name] || {}).total || 0;
              const groups = groupsOf(a.name);
              const multi = groups.length > 1;
              /* each group is a row; a subtotal row is appended when the account
                 is split across more than one remark. All share the merged
                 account/person/bank/actions cells on the left. */
              const bodyRows = groups.length ? groups.slice() : [{ remark:'', count:0, principal:0, empty:true }];
              if (multi){
                bodyRows.push({ remark:'', count:groups.reduce(function(s,g){ return s+g.count; },0),
                  principal:groups.reduce(function(s,g){ return s+g.principal; },0), isTotal:true });
              }
              const span = bodyRows.length;
              return bodyRows.map(function(g, gi){
                return (
                  <tr key={a.name + '|' + gi} className={(gi === 0 ? 'acct-start' : '') + (g.isTotal ? ' acct-total' : '') + (isFlashAccount(flash, a.name) ? ' flash' : '')}>
                    {gi === 0 ? (
                      <td className="col-actions acct-id acct-actions" rowSpan={span}>
                        <div className="stacked-btns">
                          <button type="button" className="btn ghost small" onClick={function(){ onEdit(a); }}>Edit</button>
                          <button type="button" className="btn ghost small" onClick={function(){ onTransfer(a); }}
                                  title="Move FDs to or from this account">Transfer</button>
                          <button type="button" className="btn danger-ghost small" aria-label={'Delete account ' + a.name}
                                  title={total > 0 ? 'You will be asked to move its ' + total + ' FD(s) first' : 'Delete this account'}
                                  onClick={function(){ onDelete(a); }}>Delete</button>
                        </div>
                      </td>
                    ) : null}
                    {gi === 0 ? <td className="col-acct acct-id" rowSpan={span}><span className="share-name">{a.name}</span></td> : null}
                    {gi === 0 ? <td className="acct-id" rowSpan={span}>{a.person || '—'}</td> : null}
                    {gi === 0 ? <td className="acct-id" rowSpan={span}>{a.bank || '—'}</td> : null}
                    <td>{g.isTotal ? <span className="acct-total-label">Total</span> : g.empty ? <span className="muted">—</span> : (g.remark ? g.remark : <span className="muted">no remark</span>)}</td>
                    <td className="num">{groups.length ? g.count : '0'}</td>
                    <td className="num">{groups.length ? M.fmtINR(g.principal) : '—'}{groups.length && M.fmtINRWords(g.principal) ? <div className="cell-sub cell-words">{M.fmtINRWords(g.principal)}</div> : null}</td>
                  </tr>
                );
              });
            })}
          </tbody>
        </table>
       </div>
      )}
    </div>
  );
}

function AccountForm({ initial, accounts, fdCount, busy, onCancel, onSave, readOnly, title, headNote }){
  const editing = !!initial;
  const [acc, setAcc] = useState(initial
    ? { name: initial.name, person: initial.person || '', bank: initial.bank || '' }
    : { name:'', person:'', bank:'' });
  const [err, setErr] = useState('');

  useEffect(function(){
    /* ignore ESC while the confirmation dialog is on top — it closes itself */
    function onKey(e){ if (e.key === 'Escape' && !document.querySelector('.confirm-overlay')) onCancel(); }
    document.addEventListener('keydown', onKey);
    return function(){ document.removeEventListener('keydown', onKey); };
  }, [onCancel]);

  const persons = [];
  accounts.forEach(function(a){ if (a.person && persons.indexOf(a.person) < 0) persons.push(a.person); });
  const renamed = editing && acc.name.trim() !== initial.name;

  function submit(e){
    e.preventDefault();
    if (readOnly) return;
    const name = acc.name.trim();
    if (!name){ setErr('The account needs a name.'); return; }
    const dup = accounts.some(function(a){
      return a.name.toLowerCase() === name.toLowerCase() && (!editing || a.name !== initial.name);
    });
    if (dup){ setErr('An account named “' + name + '” already exists.'); return; }
    onSave({ name:name, person:acc.person.trim(), bank:acc.bank.trim() }, editing ? initial.name : null);
  }

  return (
    <div className="overlay" onMouseDown={function(e){ if (e.target === e.currentTarget) onCancel(); }}>
      <form className="modal" style={{ maxWidth:520 }} role="dialog" aria-modal="true"
            aria-label={title || (editing ? 'Edit account' : 'Add account')} onSubmit={submit}>
        <div className="modal-head">
          <h2>{title || (editing ? 'Edit account' : 'Add an account')}</h2>
          <button type="button" className="btn ghost" aria-label="Close" onClick={onCancel}>✕</button>
        </div>
        <div className="modal-body">
          {headNote ? <div className="readonly-note">{headNote}</div> : null}
          {err && <div className="frm-err" role="alert">{err}</div>}
          <fieldset className="frm-fieldset" disabled={!!readOnly}>
          <div className="frm-section">
            <div className="frm-row">
              <Field id="acc-name" label="Account name">
                <input id="acc-name" type="text" value={acc.name} placeholder="e.g. MOM HDFC" autoFocus
                       onChange={function(e){ setAcc(Object.assign({}, acc, { name: e.target.value })); }} />
              </Field>
              <Field id="acc-person" label="Person">
                <input id="acc-person" type="text" list="acc-persons-list" value={acc.person} placeholder="e.g. Mom"
                       onChange={function(e){ setAcc(Object.assign({}, acc, { person: e.target.value })); }} />
                <datalist id="acc-persons-list">
                  {persons.map(function(p){ return <option key={p} value={p} />; })}
                </datalist>
              </Field>
            </div>
            <div className="frm-row">
              <Field id="acc-bank" label="Bank">
                <input id="acc-bank" type="text" value={acc.bank} placeholder="e.g. HDFC"
                       onChange={function(e){ setAcc(Object.assign({}, acc, { bank: e.target.value })); }} />
              </Field>
            </div>
            {renamed && fdCount > 0 ? (
              <div className="readout">
                <span>Renaming will also update <b>{fdCount}</b> FD{fdCount === 1 ? '' : 's'} under “{initial.name}”.</span>
              </div>
            ) : null}
          </div>
          </fieldset>
        </div>
        <div className="modal-foot">
          {readOnly ? (
            <button type="button" className="btn primary" onClick={onCancel}>Close</button>
          ) : (
            <React.Fragment>
              <button type="button" className="btn" onClick={onCancel}>Cancel</button>
              <button type="submit" className="btn primary" disabled={busy}>
                {busy ? 'Saving…' : (editing ? 'Save changes' : 'Add account')}
              </button>
            </React.Fragment>
          )}
        </div>
      </form>
    </div>
  );
}

/* ============================== add / edit form ============================== */
function Field({ id, label, auto, children }){
  return (
    <div className="field">
      <label htmlFor={id}>{label}{auto ? <span className="auto-badge">auto</span> : null}</label>
      {children}
    </div>
  );
}

function FdForm({ initial, accounts, busy, onCancel, onSave, readOnly, title, headNote }){
  const editing = !!initial;
  const [st, setSt] = useState(function(){
    if (!initial) return { f: M.emptyForm(), pins: {} };
    const x = M.formFromFd(initial);
    return { f: M.solveForm(x.f, x.pins), pins: x.pins };
  });
  const [newAcc, setNewAcc] = useState(null);
  const [err, setErr] = useState('');
  const f = st.f, pins = st.pins;

  useEffect(function(){
    /* ignore ESC while the confirmation dialog is on top — it closes itself */
    function onKey(e){ if (e.key === 'Escape' && !document.querySelector('.confirm-overlay')) onCancel(); }
    document.addEventListener('keydown', onKey);
    document.body.style.overflow = 'hidden';
    return function(){
      document.removeEventListener('keydown', onKey);
      document.body.style.overflow = '';
    };
  }, [onCancel]);

  function pinKeyFor(name){
    if (name === 'start' || name === 'end' || name === 'rate' || name === 'principal' || name === 'maturity') return name;
    if (name === 'ty' || name === 'tm' || name === 'td') return 'tenure';
    return null;
  }

  function setField(name, value){
    setSt(function(prev){
      const nf = Object.assign({}, prev.f);
      nf[name] = value;
      const np = Object.assign({}, prev.pins);
      const pk = pinKeyFor(name);
      if (pk === 'tenure') np.tenure = M.tenureEntered(nf);
      else if (pk) np[pk] = value !== '';
      if (name === 'interestOnMaturity') np.maturity = false; /* payout toggle un-pins maturity */
      if (name === 'status') nf.showInDashboard = value === 'Active'; /* inactive never counts */
      return { f: M.solveForm(nf, np), pins: np };
    });
  }

  /* "↻ Recalculate" — un-pin maturity, treat rate+principal as the inputs */
  function recalcMaturity(){
    setSt(function(prev){
      const np = Object.assign({}, prev.pins, {
        maturity: false,
        rate: prev.f.rate !== '',
        principal: prev.f.principal !== ''
      });
      return { f: M.solveForm(Object.assign({}, prev.f), np), pins: np };
    });
  }

  const isAuto = function(field, val){ return val !== '' && !pins[field]; };
  const tDays = M.formDays(f);
  const tenLbl = M.tenureEntered(f) ? M.tenureLabel(M.tenureOf(f)) : (tDays ? null : '—');
  const liveTen = tenLbl !== null ? tenLbl : M.tenureLabel(M.diffYMD(f.start, f.end));
  const P = parseFloat(f.principal), Mx = parseFloat(f.maturity);
  const liveInterest = (!isNaN(P) && !isNaN(Mx)) ? Mx - P : null;

  const persons = [];
  accounts.forEach(function(a){ if (a.person && persons.indexOf(a.person) < 0) persons.push(a.person); });
  const byPerson = {};
  accounts.forEach(function(a){
    const k = a.person || 'Other';
    (byPerson[k] = byPerson[k] || []).push(a);
  });

  function submit(e){
    e.preventDefault();
    if (readOnly) return;
    const account = newAcc ? String(newAcc.name || '').trim() : f.account;
    if (!account){ setErr('Pick an account, or add a new one.'); return; }
    if (M.parseISO(f.start) && M.parseISO(f.end) && M.diffDays(f.start, f.end) < 0){
      setErr('The end date is before the start date.'); return;
    }
    const bad = ['rate', 'principal', 'maturity'].find(function(k){ return f[k] !== '' && isNaN(parseFloat(f[k])); });
    if (bad){ setErr('“' + bad + '” must be a number.'); return; }
    const fd = {
      account: account,
      accountNumber: String(f.accountNumber).trim(),
      startDate: f.start, endDate: f.end,
      tenureDays: tDays === null ? '' : tDays,
      rate: f.rate === '' ? '' : parseFloat(f.rate),
      principal: f.principal === '' ? '' : parseFloat(f.principal),
      /* periodic payout: maturity is always stored as the principal */
      maturity: !f.interestOnMaturity
        ? (f.principal === '' ? '' : parseFloat(f.principal))
        : (f.maturity === '' ? '' : parseFloat(f.maturity)),
      interestOnMaturity: f.interestOnMaturity,
      compounding: f.compounding,
      status: f.status,
      showInDashboard: f.status === 'Active' ? f.showInDashboard : false,
      remarks: String(f.remarks).trim()
    };
    if (editing) fd.id = initial.id;
    onSave(fd, newAcc ? {
      name: account,
      person: String(newAcc.person || '').trim(),
      bank: String(newAcc.bank || '').trim()
    } : null);
  }

  return (
    <div className="overlay" onMouseDown={function(e){ if (e.target === e.currentTarget) onCancel(); }}>
      <form className="modal" role="dialog" aria-modal="true" aria-label={title || (editing ? 'Edit FD' : 'Add FD')} onSubmit={submit}>
        <div className="modal-head">
          <h2>{title || (editing ? 'Edit FD #' + initial.id : 'Add a fixed deposit')}</h2>
          <button type="button" className="btn ghost" aria-label="Close" onClick={onCancel}>✕</button>
        </div>
        <div className="modal-body">
          {headNote ? <div className="readonly-note">{headNote}</div> : null}
          {err && <div className="frm-err" role="alert">{err}</div>}

          <fieldset className="frm-fieldset" disabled={!!readOnly}>
          <div className="frm-section">
            <div className="frm-section-title">Account</div>
            {!newAcc ? (
              <div className="frm-row">
                <Field id="f-account" label="Account">
                  <select id="f-account" value={f.account}
                          onChange={function(e){
                            if (e.target.value === '__new__'){ setNewAcc({ name:'', person:'', bank:'' }); }
                            else setField('account', e.target.value);
                          }}>
                    <option value="">— choose —</option>
                    {Object.keys(byPerson).sort().map(function(p){
                      return (
                        <optgroup key={p} label={p}>
                          {byPerson[p].map(function(a){ return <option key={a.name} value={a.name}>{a.name}</option>; })}
                        </optgroup>
                      );
                    })}
                    <option value="__new__">＋ Add new account…</option>
                  </select>
                </Field>
                <Field id="f-acctno" label="FD account number">
                  <input id="f-acctno" type="text" inputMode="numeric" pattern="[0-9]*" value={f.accountNumber} placeholder="digits only — enter 0 if N/A"
                         onChange={function(e){ setField('accountNumber', e.target.value.replace(/[^0-9]/g, '')); }} />
                </Field>
              </div>
            ) : (
              <div>
                <div className="frm-row three">
                  <Field id="na-name" label="New account name">
                    <input id="na-name" type="text" value={newAcc.name} placeholder="e.g. MOM HDFC" autoFocus
                           onChange={function(e){ setNewAcc(Object.assign({}, newAcc, { name: e.target.value })); }} />
                  </Field>
                  <Field id="na-person" label="Person">
                    <input id="na-person" type="text" list="persons-list" value={newAcc.person} placeholder="e.g. Mom"
                           onChange={function(e){ setNewAcc(Object.assign({}, newAcc, { person: e.target.value })); }} />
                    <datalist id="persons-list">
                      {persons.map(function(p){ return <option key={p} value={p} />; })}
                    </datalist>
                  </Field>
                  <Field id="na-bank" label="Bank">
                    <input id="na-bank" type="text" value={newAcc.bank} placeholder="e.g. HDFC"
                           onChange={function(e){ setNewAcc(Object.assign({}, newAcc, { bank: e.target.value })); }} />
                  </Field>
                </div>
                <div className="frm-row">
                  <div>
                    <button type="button" className="link-btn" onClick={function(){ setNewAcc(null); }}>
                      ← use an existing account instead
                    </button>
                  </div>
                  <Field id="f-acctno2" label="FD account number">
                    <input id="f-acctno2" type="text" inputMode="numeric" pattern="[0-9]*" value={f.accountNumber} placeholder="digits only — enter 0 if N/A"
                           onChange={function(e){ setField('accountNumber', e.target.value.replace(/[^0-9]/g, '')); }} />
                  </Field>
                </div>
              </div>
            )}
          </div>

          <div className="frm-section">
            <div className="frm-section-title">Dates &amp; tenure — fill any two, the third is computed</div>
            <div className="frm-row three">
              <Field id="f-start" label="Start date" auto={isAuto('start', f.start)}>
                <input id="f-start" type="date" value={f.start}
                       onChange={function(e){ setField('start', e.target.value); }} />
              </Field>
              <Field id="f-end" label="End date" auto={isAuto('end', f.end)}>
                <input id="f-end" type="date" value={f.end}
                       onChange={function(e){ setField('end', e.target.value); }} />
              </Field>
              <Field id="f-ty" label="Tenure" auto={isAuto('tenure', M.tenureEntered(f) ? 'x' : '')}>
                <div className="ten-grid">
                  <span className="ten-cell">
                    <input id="f-ty" aria-label="Tenure years" inputMode="numeric" value={f.ty}
                           onChange={function(e){ setField('ty', e.target.value.replace(/[^0-9]/g, '')); }} />
                    <span className="ten-unit">y</span>
                  </span>
                  <span className="ten-cell">
                    <input aria-label="Tenure months" inputMode="numeric" value={f.tm}
                           onChange={function(e){ setField('tm', e.target.value.replace(/[^0-9]/g, '')); }} />
                    <span className="ten-unit">m</span>
                  </span>
                  <span className="ten-cell">
                    <input aria-label="Tenure days" inputMode="numeric" value={f.td}
                           onChange={function(e){ setField('td', e.target.value.replace(/[^0-9]/g, '')); }} />
                    <span className="ten-unit">d</span>
                  </span>
                </div>
              </Field>
            </div>
            <div className="readout">
              <span>Tenure&nbsp;<b>{liveTen}</b>{tDays !== null ? <span className="muted"> · {tDays} days</span> : null}</span>
            </div>
          </div>

          <div className="frm-section">
            <div className="frm-section-title">Interest &amp; amounts — fill any two, the third is computed</div>
            <div className="frm-row">
              <div className="field">
                <label>Interest payout</label>
                <div className="seg" role="group" aria-label="Interest payout">
                  <button type="button" className={'seg-btn' + (f.interestOnMaturity ? ' on' : '')}
                          onClick={function(){ if (!f.interestOnMaturity) setField('interestOnMaturity', true); }}>
                    On maturity
                  </button>
                  <button type="button" className={'seg-btn' + (!f.interestOnMaturity ? ' on' : '')}
                          onClick={function(){ if (f.interestOnMaturity) setField('interestOnMaturity', false); }}>
                    Periodic payout
                  </button>
                </div>
              </div>
              <Field id="f-comp" label="Compounding">
                <select id="f-comp" value={f.compounding}
                        onChange={function(e){ setField('compounding', e.target.value); }}>
                  <option value="quarterly">Quarterly (most banks)</option>
                  <option value="monthly">Monthly</option>
                  <option value="half-yearly">Half-yearly</option>
                  <option value="annually">Annually</option>
                  <option value="simple">Simple interest</option>
                </select>
              </Field>
            </div>
            <div className="frm-row three">
              <Field id="f-rate" label="Interest % p.a." auto={isAuto('rate', f.rate)}>
                <input id="f-rate" className="num-in" inputMode="decimal" value={f.rate} placeholder="7.1"
                       onChange={function(e){ setField('rate', e.target.value); }} />
              </Field>
              <Field id="f-principal" label="Principal (₹)" auto={isAuto('principal', f.principal)}>
                <input id="f-principal" className="num-in" inputMode="decimal" value={f.principal} placeholder="50000"
                       onChange={function(e){ setField('principal', e.target.value); }} />
                {M.fmtINRWords(f.principal) ? <div className="muted small num-words">{M.fmtINRWords(f.principal)}</div> : null}
              </Field>
              <Field id="f-maturity" label="Maturity (₹)" auto={isAuto('maturity', f.maturity)}>
                <input id="f-maturity" className="num-in" inputMode="decimal" value={f.maturity}
                       disabled={!f.interestOnMaturity}
                       title={!f.interestOnMaturity ? 'Periodic payout: maturity always equals the principal' : ''}
                       onChange={function(e){ setField('maturity', e.target.value); }} />
                {M.fmtINRWords(f.maturity) ? <div className="muted small num-words">{M.fmtINRWords(f.maturity)}</div> : null}
              </Field>
            </div>
            <div className="readout">
              <span>Interest&nbsp;<b className="accent">{liveInterest === null ? '—' : M.fmtINR(M.round2(liveInterest))}</b>{liveInterest !== null && M.fmtINRWords(liveInterest) ? <span className="muted"> · {M.fmtINRWords(M.round2(liveInterest))}</span> : null}<span className="muted"> (maturity − principal)</span></span>
              {!f.interestOnMaturity ? <span className="muted">Periodic payout: maturity is locked to the principal; the rate is informational.</span> : null}
            </div>
            <div style={{ marginBottom: 11 }}>
              <button type="button" className="link-btn" onClick={recalcMaturity} disabled={!f.interestOnMaturity}>
                ↻ Recalculate maturity from principal &amp; rate
              </button>
              <span className="muted small"> — use after correcting a rate; overrides the typed maturity</span>
            </div>
          </div>

          <div className="frm-section">
            <div className="frm-section-title">Status</div>
            <div className="frm-row three">
              <Field id="f-status" label="Status">
                <select id="f-status" value={f.status} onChange={function(e){ setField('status', e.target.value); }}>
                  <option value="Active">Active</option>
                  <option value="Inactive">Inactive</option>
                </select>
              </Field>
              <div className="field">
                <label htmlFor="f-show">Dashboard</label>
                <div className="check-row">
                  <input id="f-show" type="checkbox" checked={f.showInDashboard && f.status === 'Active'}
                         disabled={f.status !== 'Active'}
                         onChange={function(e){ setField('showInDashboard', e.target.checked); }} />
                  <label htmlFor="f-show" style={{ margin:0, fontWeight:500 }}>Count in dashboard totals</label>
                </div>
                {f.status !== 'Active' ? <div className="muted small">Inactive FDs are never counted in the dashboard.</div> : null}
              </div>
              <Field id="f-remarks" label="Remarks">
                <input id="f-remarks" type="text" value={f.remarks} placeholder="nominee, sub-bank, notes…"
                       onChange={function(e){ setField('remarks', e.target.value); }} />
              </Field>
            </div>
          </div>
          </fieldset>
        </div>
        <div className="modal-foot">
          {readOnly ? (
            <button type="button" className="btn primary" onClick={onCancel}>Close</button>
          ) : (
            <React.Fragment>
              <button type="button" className="btn" onClick={onCancel}>Cancel</button>
              <button type="submit" className="btn primary" disabled={busy}>
                {busy ? 'Saving…' : (editing ? 'Save changes' : 'Add FD')}
              </button>
            </React.Fragment>
          )}
        </div>
      </form>
    </div>
  );
}

/* ============================== transfer FDs ============================== */
/* Move every FD from one account to another — out of this account, or into it
 * from another. Used on its own (Transfer button) and as the first step of
 * deleting an account that still holds FDs. */
function TransferDialog({ account, accounts, fds, forceOut, busy, onCancel, onSubmit }){
  const [dir, setDir] = useState('out');         /* 'out' = this→other, 'in' = other→this */
  const [other, setOther] = useState('');
  const [checked, setChecked] = useState({});
  const [err, setErr] = useState('');
  const effDir = forceOut ? 'out' : dir;
  const low = function(s){ return String(s == null ? '' : s).toLowerCase(); };

  const countOf = function(name){ return fds.filter(function(f){ return low(f.account) === low(name); }).length; };
  const others = accounts.filter(function(a){ return a.name !== account.name; });
  /* the account whose FDs you pick from */
  const fromAccount = effDir === 'out' ? account.name : other;
  const toAccount = effDir === 'out' ? other : account.name;
  const candidates = fds.filter(function(f){ return fromAccount && low(f.account) === low(fromAccount); });
  const selectedIds = candidates.filter(function(f){ return checked[f.id]; }).map(function(f){ return f.id; });

  /* tick every FD of the source account by default whenever that set changes */
  useEffect(function(){
    const init = {};
    fds.forEach(function(f){ if (fromAccount && low(f.account) === low(fromAccount)) init[f.id] = true; });
    setChecked(init);
    setErr('');
  }, [effDir, other]);

  useEffect(function(){
    function onKey(e){ if (e.key === 'Escape' && !document.querySelector('.confirm-overlay')) onCancel(); }
    document.addEventListener('keydown', onKey);
    return function(){ document.removeEventListener('keydown', onKey); };
  }, [onCancel]);

  function toggle(id){ setChecked(function(c){ const n = Object.assign({}, c); n[id] = !n[id]; return n; }); }
  function setAll(val){ const n = {}; candidates.forEach(function(f){ n[f.id] = val; }); setChecked(n); }

  function submit(e){
    e.preventDefault();
    if (!other){ setErr('Choose the other account first.'); return; }
    if (selectedIds.length === 0){ setErr('Tick at least one FD to move.'); return; }
    onSubmit(selectedIds, fromAccount, toAccount, selectedIds.length);
  }

  return (
    <div className="overlay" onMouseDown={function(e){ if (e.target === e.currentTarget) onCancel(); }}>
      <form className="modal" style={{ maxWidth:560 }} role="dialog" aria-modal="true"
            aria-label="Transfer FDs" onSubmit={submit}>
        <div className="modal-head">
          <h2>{forceOut ? 'Move FDs out of “' + account.name + '”' : 'Transfer FDs'}</h2>
          <button type="button" className="btn ghost" aria-label="Close" onClick={onCancel}>✕</button>
        </div>
        <div className="modal-body">
          {forceOut ? (
            <div className="readonly-note">“{account.name}” still holds {countOf(account.name)} FD{countOf(account.name) === 1 ? '' : 's'}. Pick the ones to move to another account, then you can delete it.</div>
          ) : null}
          {err && <div className="frm-err" role="alert">{err}</div>}
          {others.length === 0 ? (
            <div className="panel-empty" style={{ padding:'10px 0' }}>There's no other account to transfer with — add one first.</div>
          ) : (
          <div className="frm-section">
            {!forceOut ? (
              <div className="field" style={{ marginBottom:12 }}>
                <label>Direction</label>
                <div className="seg" role="group" aria-label="Transfer direction">
                  <button type="button" className={'seg-btn' + (dir === 'out' ? ' on' : '')}
                          onClick={function(){ setDir('out'); }}>Move out of this</button>
                  <button type="button" className={'seg-btn' + (dir === 'in' ? ' on' : '')}
                          onClick={function(){ setDir('in'); }}>Move into this</button>
                </div>
              </div>
            ) : null}
            <div className="frm-row">
              <Field id="tr-other" label={effDir === 'out' ? 'Move selected FDs to' : 'Take selected FDs from'}>
                <select id="tr-other" value={other} onChange={function(e){ setOther(e.target.value); }}>
                  <option value="">— choose account —</option>
                  {others.map(function(a){
                    return <option key={a.name} value={a.name}>{a.name} ({countOf(a.name)} FD{countOf(a.name) === 1 ? '' : 's'})</option>;
                  })}
                </select>
              </Field>
            </div>
            {!fromAccount ? (
              <div className="panel-empty" style={{ padding:'6px 0' }}>Choose the account above to list its FDs.</div>
            ) : candidates.length === 0 ? (
              <div className="panel-empty" style={{ padding:'6px 0' }}>“{fromAccount}” has no FDs to move.</div>
            ) : (
              <div>
                <div className="tr-tools">
                  <span className="muted small">{selectedIds.length} of {candidates.length} selected</span>
                  <button type="button" className="link-btn" onClick={function(){ setAll(true); }}>Select all</button>
                  <button type="button" className="link-btn" onClick={function(){ setAll(false); }}>Clear</button>
                </div>
                <div className="tr-list">
                  {candidates.map(function(fd){
                    const intr = (fd.maturity > 0) ? fd.maturity - fd.principal : null;
                    return (
                      <label key={fd.id} className="tr-fd">
                        <input type="checkbox" checked={!!checked[fd.id]} onChange={function(){ toggle(fd.id); }} />
                        <span className="tr-fd-main">
                          <span className="tr-fd-top">{fd.accountNumber ? fd.accountNumber : 'FD #' + fd.id} · {M.fmtINR(fd.principal || '')} → {M.fmtINR(fd.maturity || '')}</span>
                          <span className="tr-fd-sub muted small">
                            {M.fmtRate(fd.rate)} · {M.fmtDate(fd.startDate)} – {M.fmtDate(fd.endDate)} · {fd.status}
                            {intr !== null ? ' · interest ' + M.fmtINR(M.round2(intr)) : ''}
                            {fd.remarks ? ' · ' + fd.remarks : ''}
                          </span>
                        </span>
                      </label>
                    );
                  })}
                </div>
                <div className="readout">
                  <span>{other
                    ? 'Moves ' + selectedIds.length + ' of ' + candidates.length + ' FD' + (candidates.length === 1 ? '' : 's') + ' from “' + fromAccount + '” to “' + toAccount + '”.'
                    : 'Pick the other account above to finish.'}</span>
                </div>
              </div>
            )}
          </div>
          )}
        </div>
        <div className="modal-foot">
          <button type="button" className="btn" onClick={onCancel}>Cancel</button>
          <button type="submit" className="btn primary" disabled={busy || others.length === 0}>
            {busy ? 'Working…' : 'Review move'}
          </button>
        </div>
      </form>
    </div>
  );
}

/* ============================== confirmation dialog ============================== */
/* One reusable yes/no step for add / edit / delete. Cancel is the focused,
 * safe default, so a stray Enter or mis-tap never destroys anything. */
function ConfirmDialog({ title, lines, confirmLabel, danger, busy, error, onConfirm, onCancel }){
  useEffect(function(){
    function onKey(e){ if (e.key === 'Escape') onCancel(); }
    document.addEventListener('keydown', onKey);
    return function(){ document.removeEventListener('keydown', onKey); };
  }, [onCancel]);
  return (
    <div className="overlay confirm-overlay" onMouseDown={function(e){ if (e.target === e.currentTarget) onCancel(); }}>
      <div className="modal confirm-modal" role="alertdialog" aria-modal="true" aria-label={title}>
        <div className="modal-head"><h2>{title}</h2></div>
        <div className="modal-body">
          <div className="confirm-summary">
            {lines.filter(Boolean).map(function(l, i){
              return <div key={i} className={i === 0 ? '' : 'muted small'} style={i === 0 ? { fontWeight:650 } : undefined}>{l}</div>;
            })}
          </div>
          {error ? <div className="frm-err confirm-err" role="alert">{error}</div> : null}
        </div>
        <div className="modal-foot">
          <button type="button" className="btn" autoFocus onClick={onCancel}>Cancel</button>
          <button type="button" className={'btn ' + (danger ? 'danger-solid' : 'primary')}
                  onClick={onConfirm} disabled={busy}>
            {busy ? 'Please wait…' : confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================== app shell ============================== */
function App(){
  const TABS = { dashboard:1, all:1, accounts:1, history:1, reminders: LIVE ? 1 : 0 };
  const [auth, setAuth] = useState(null);
  const [data, setData] = useState(null);
  const [tab, setTab] = useState(function(){ const t = FDPrefs.getTab(); return TABS[t] ? t : 'dashboard'; });
  const [allKey, setAllKey] = useState(0); /* bump to remount All FDs → reset its filters */
  const [editing, setEditing] = useState(null);      /* null | 'new' | fd */
  const [accEditing, setAccEditing] = useState(null); /* null | 'new' | account */
  const [transferring, setTransferring] = useState(null); /* {account, forceOut} */
  const [confirm, setConfirm] = useState(null);      /* {title, lines, confirmLabel, danger, run, error} */
  const confirmRef = useRef(null);                    /* the dialog open right now (read after an await) */
  useEffect(function(){ confirmRef.current = confirm; }, [confirm]);
  const saveErrRef = useRef('');                      /* why the last save failed */
  const [viewing, setViewing] = useState(null);      /* a log entry being inspected read-only */
  /* ---- History month + data freshness --------------------------------------
   * want = the month History should show (month null = "this month", the
   * backend's default); every change of it bumps want.seq. A reply counts for
   * History only if no newer want was made after its request went out.
   * loadedSeq = the newest want a reply has satisfied, so History is loading
   * while loadedSeq < want.seq — including "back to this month". */
  const wantRef = useRef({ month: null, seq: 0 });
  const [want, setWant] = useState(wantRef.current);
  const loadedSeqRef = useRef(0);
  const [loadedSeq, setLoadedSeq] = useState(0);
  const shownWantRef = useRef(null);  /* the want.month whose log is on screen (null = this month) */
  const lastSyncRef = useRef(0);      /* when the last reply arrived (ms) */
  function wantMonth(m){
    const w = { month: m || null, seq: wantRef.current.seq + 1 };
    wantRef.current = w; setWant(w);
    return w;
  }
  function markLoaded(seq){
    if (seq > loadedSeqRef.current){ loadedSeqRef.current = seq; setLoadedSeq(seq); }
  }
  /* History is showing, or already fetching, this month: going "back to this
   * month" again must not start another load (that would only discard the
   * reply already on its way and make the wait longer). */
  function wantsThisMonth(){ return wantRef.current.month === null; }
  function withLogMonth(body){
    const m = wantRef.current.month;
    return m ? Object.assign({}, body, { logMonth: m }) : body;
  }
  /* The request that asked for a month failed: unless another reply has shown
   * it meanwhile, go back to the month that IS on screen, so History never
   * sits on "Loading…" or keeps its controls aimed at a month it can't show. */
  function rollBackWant(asked){
    if (wantRef.current.seq === asked.seq && loadedSeqRef.current < asked.seq){
      markLoaded(wantMonth(shownWantRef.current).seq);
    }
  }
  /* Apply a backend reply. `asked` = the want ({month, seq}) the request went
   * out with: if a newer want was made since, keep the log on screen (that
   * want's own reply is on its way) but take the fresh FDs etc.
   * logMonths === null marks an older backend that predates month browsing;
   * History then shows the latest-15 list it sends. */
  function applyServer(j, asked){
    const fresh = asked.seq === wantRef.current.seq;
    if (fresh){ shownWantRef.current = asked.month; markLoaded(asked.seq); }
    lastSyncRef.current = Date.now();
    setData(function(prev){
      const keep = !!prev && !fresh;
      return {
        fds: j.fds, accounts: j.accounts,
        log: keep ? prev.log : (j.log || []),
        logMonth: keep ? prev.logMonth : (j.logMonth || null),
        logMonths: keep ? prev.logMonths : (Array.isArray(j.logMonths) ? j.logMonths : null),
        vapidPublic: j.vapidPublic || '', pushSubs: j.pushSubs || []
      };
    });
  }
  /* Demo (no network): slice the in-memory log for the month wanted. */
  function demoLogView(d){
    const v = logViewOf(d._rawLog, wantRef.current.month, thisMonthKey(), d.fds, d.accounts);
    shownWantRef.current = wantRef.current.month;
    return Object.assign({}, d, v);
  }
  function setDemo(d){ shownWantRef.current = wantRef.current.month; markLoaded(wantRef.current.seq); setData(d); }

  /* Requests in flight, for the top bar and the "Saving…" / "Loading…" pill.
   * Mirrored in a ref so code finishing after an await sees the live count. */
  const inflightRef = useRef({ all: 0, saves: 0 });
  const [inflight, setInflight] = useState(inflightRef.current);
  function countInflight(dAll, dSaves){
    const p = inflightRef.current;
    inflightRef.current = { all: p.all + dAll, saves: p.saves + dSaves };
    setInflight(inflightRef.current);
  }
  async function call(body, isSave){
    countInflight(1, isSave ? 1 : 0);
    try { return await postApi(body); }
    finally { countInflight(-1, isSave ? -1 : 0); }
  }

  /* After a save, the changed FD / account rows glow for a few seconds. */
  const [flash, setFlash] = useState(null); /* {fds:[ids], account:name|null, n} */
  useEffect(function(){
    if (!flash) return;
    const t = setTimeout(function(){ setFlash(null); }, 5000);
    return function(){ clearTimeout(t); };
  }, [flash]);
  const [busy, setBusy] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [loginErr, setLoginErr] = useState('');
  const [toast, setToast] = useState(null); /* {msg, kind:'success'|'error'} */
  const [theme, setThemeState] = useState(function(){ return FDPrefs.getTheme(); });
  const [pushState, setPushState] = useState('off'); /* 'unsupported'|'off'|'on'|'denied'|'busy' */
  const [thisEndpoint, setThisEndpoint] = useState(''); /* this device's push endpoint, if subscribed */
  /* true while trying the sign-in remembered on this device */
  const [boot, setBoot] = useState(function(){ return !!FDPrefs.getAuth(); });

  useEffect(function(){
    if (!toast) return;
    const t = setTimeout(function(){ setToast(null); }, 5000); /* self-hides in 5s */
    return function(){ clearTimeout(t); };
  }, [toast]);

  /* Auto sign-in with the cookie saved on this device, once, on load. */
  useEffect(function(){
    const pw = FDPrefs.getAuth();
    if (!pw){ setBoot(false); return; }
    if (!LIVE){
      setDemo(demoSeed());
      setAuth({ password: pw });
      setPushAuth(pw);
      setBoot(false);
      return;
    }
    (async function(){
      try {
        const asked = wantRef.current;
        const j = await call(withLogMonth({ action:'list', password: pw }));
        applyServer(j, asked);
        setAuth({ password: pw });
        setPushAuth(pw);
      } catch (e) {
        if (e.message === 'unauthorized'){
          FDPrefs.clearAuth();
          setLoginErr('The saved sign-in is no longer valid — enter the family password.');
        } else {
          setLoginErr('Could not sign in automatically: ' + e.message);
        }
      } finally {
        setBoot(false);
      }
    })();
  }, []);

  function notifyOk(msg, fromSave){ setToast({ msg: msg, kind: 'success', save: !!fromSave }); }
  function notifyError(msg){ setToast({ msg: msg, kind: 'error' }); }

  /* reflect the current reminder state + this device's endpoint once signed in */
  useEffect(function(){
    if (!auth) return;
    thisDeviceEndpoint().then(function(ep){ setThisEndpoint(ep); }).catch(function(){});
    if (!LIVE || !PUSH_SUPPORTED){ setPushState('unsupported'); return; }
    if (typeof Notification !== 'undefined' && Notification.permission === 'denied'){ setPushState('denied'); return; }
    getPushSubscription().then(function(sub){ setPushState(sub ? 'on' : 'off'); }).catch(function(){ setPushState('off'); });
  }, [auth]);

  function applyPushSubs(j){ if (j && j.pushSubs) setData(function(d){ return Object.assign({}, d, { pushSubs: j.pushSubs }); }); }

  async function enableReminders(){
    setPushState('busy');
    try {
      const j = await enablePush(data && data.vapidPublic, FDPrefs.getDevice());
      applyPushSubs(j);
      setThisEndpoint(await thisDeviceEndpoint());
      setPushState('on');
      notifyOk('Maturity reminders on for this device');
    } catch (e) {
      setPushState(typeof Notification !== 'undefined' && Notification.permission === 'denied' ? 'denied' : 'off');
      notifyError(e.message);
    }
  }
  async function disableReminders(){
    setPushState('busy');
    try {
      const j = await disablePush();
      applyPushSubs(j);
      setThisEndpoint('');
      notifyOk('Reminders turned off on this device');
    } catch (e) {
      notifyError('Could not turn off: ' + e.message);
    } finally {
      setPushState('off');
    }
  }
  async function unsubscribeDevice(s){
    try {
      const j = await unsubscribeEndpoint(s.endpoint);
      applyPushSubs(j);
      if (thisEndpoint && s.endpoint === thisEndpoint){ setThisEndpoint(''); setPushState('off'); }
      notifyOk('Removed ' + (s.device ? '“' + s.device + '”' : 'that device'));
    } catch (e) {
      notifyError('Could not remove: ' + e.message);
    }
  }

  function toggleTheme(){
    const next = theme === 'dark' ? 'light' : 'dark';
    setThemeState(next);
    FDPrefs.setTheme(next);
  }

  async function login(pw, deviceName){
    setLoginErr('');
    if (!pw){ setLoginErr('Enter the family password.'); return; }
    if (!LIVE){
      FDPrefs.setAuth(pw);
      FDPrefs.setDevice(deviceName || '');
      setDemo(demoSeed());
      setAuth({ password: pw });
      setPushAuth(pw);
      return;
    }
    setBusy(true);
    try {
      const asked = wantRef.current;
      const j = await call(withLogMonth({ action:'list', password: pw }));
      FDPrefs.setAuth(pw); /* remember on this device until Sign out */
      FDPrefs.setDevice(deviceName || '');
      applyServer(j, asked);
      setAuth({ password: pw });
      setPushAuth(pw);
    } catch (e) {
      setLoginErr(e.message === 'unauthorized' ? 'Wrong password — try again.' : 'Could not open the ledger: ' + e.message);
    } finally {
      setBusy(false);
    }
  }

  function logout(){
    FDPrefs.clearAuth();
    markLoaded(wantMonth(null).seq); shownWantRef.current = null; lastSyncRef.current = 0;
    setAuth(null);
    setData(null);
    setTab('dashboard');
    setLoginErr('');
    setToast(null);
  }

  /* Every save goes through here (always from a confirm dialog), and stays
   * honest: nothing changes on screen until the sheet has accepted it.
   * Meanwhile the top bar runs and a pill says "Saving…"; afterwards a toast
   * confirms what was saved and the changed row glows. If it fails, the reason
   * is shown inside that dialog (see runConfirm) and false is returned.
   * opts: { saved: toast text, flash: {fds, account} } to override either. */
  const FAIL_PREFIX = { undo: 'Couldn’t undo', delete: 'Not deleted', accountDelete: 'Not deleted', transferFds: 'Not moved' };
  async function mutate(body, opts){
    setBusy(true);
    setToast(null); /* an older "✓ …" must not sit where this save's "Saving…" goes */
    try {
      const fullBody = Object.assign({}, body, { device: FDPrefs.getDevice() });
      const prevFds = data ? data.fds : [];
      let nextFds;
      if (LIVE){
        const asked = wantRef.current;
        const j = await call(withLogMonth(Object.assign({}, fullBody, { password: auth.password })), true);
        applyServer(j, asked);
        nextFds = j.fds;
      } else {
        const next = demoMutate(data, withLogMonth(fullBody)); /* throws on guard violations, like the backend */
        setDemo(next);
        nextFds = next.fds;
      }
      notifyOk(afterSave(body, prevFds, nextFds, opts), true);
      return true;
    } catch (e) {
      saveErrRef.current = e.message === 'unauthorized'
        ? 'Session rejected — the password may have changed. Reload the page and sign in again.'
        : (FAIL_PREFIX[body.action] || 'Not saved') + ': ' + e.message;
      return false;
    } finally {
      setBusy(false);
    }
  }

  /* After a save: highlight what changed and return the confirmation text. */
  function afterSave(body, prevFds, nextFds, opts){
    opts = opts || {};
    const a = body.action;
    let ids = [], account = null, newId = null;
    if (a === 'create'){
      const had = {};
      (prevFds || []).forEach(function(f){ had[f.id] = true; });
      const added = (nextFds || []).filter(function(f){ return !had[f.id]; }).map(function(f){ return f.id; });
      if (added.length){ newId = Math.max.apply(null, added); ids = [newId]; }
    } else if (a === 'update') ids = [body.fd && body.fd.id];
    else if (a === 'markDone') ids = [body.id];
    else if (a === 'transferFds') ids = (body.ids || []).slice();
    else if (a === 'accountCreate' || a === 'accountUpdate') account = body.account && body.account.name;
    if (body.newAccount && body.newAccount.name) account = body.newAccount.name;
    const f = opts.flash || { fds: ids.filter(function(x){ return x != null; }), account: account };
    if ((f.fds && f.fds.length) || f.account) setFlash({ fds: f.fds || [], account: f.account || null, n: Date.now() });
    if (opts.saved) return opts.saved;
    const n = (body.ids || []).length;
    return ({
      create: newId ? 'Saved — FD #' + newId + ' added' : 'Saved — FD added',
      update: 'Saved — changes to FD #' + (body.fd && body.fd.id),
      delete: 'Deleted FD #' + body.id,
      markDone: 'FD #' + body.id + ' marked as done — moved to Inactive',
      accountCreate: 'Saved — account “' + (body.account && body.account.name) + '” added',
      accountUpdate: 'Saved — account “' + (body.account && body.account.name) + '” updated',
      accountDelete: 'Deleted account “' + body.name + '”',
      transferFds: 'Moved ' + n + ' FD' + (n === 1 ? '' : 's') + ' to “' + body.to + '”',
      undo: 'Change undone'
    })[a] || 'Saved';
  }

  /* Every add / edit / delete passes through one styled yes/no dialog. On
   * success it closes; on failure it stays open and says why, right where the
   * parent is looking (or, if they closed it meanwhile, a toast does). */
  async function runConfirm(){
    const mine = confirm;
    const same = function(c){ return !!c && c.run === mine.run; };
    if (mine.error) setConfirm(function(c){ return same(c) ? Object.assign({}, c, { error: '' }) : c; });
    const ok = await mine.run();
    if (ok){ setConfirm(function(c){ return same(c) ? null : c; }); return; }
    const msg = saveErrRef.current;
    if (same(confirmRef.current)) setConfirm(function(c){ return same(c) ? Object.assign({}, c, { error: msg }) : c; });
    else notifyError(msg);
  }

  function fdSummary(fd){
    return [
      fd.account + (fd.accountNumber ? ' · ' + fd.accountNumber : ''),
      M.fmtINR(fd.principal === '' ? '' : fd.principal) + ' → ' + M.fmtINR(fd.maturity === '' ? '' : fd.maturity) +
        (fd.rate ? ' at ' + M.fmtRate(fd.rate) : ''),
      (fd.startDate || fd.endDate) ? (M.fmtDate(fd.startDate) + ' to ' + M.fmtDate(fd.endDate)) : ''
    ];
  }

  function handleSave(fd, newAccount){
    setConfirm({
      title: fd.id ? 'Save changes to FD #' + fd.id + '?' : 'Add this FD?',
      lines: fdSummary(fd).concat(newAccount ? ['Also adds new account “' + newAccount.name + '”.'] : []),
      confirmLabel: fd.id ? 'Yes, save changes' : 'Yes, add FD',
      danger: false,
      run: async function(){
        const ok = await mutate({ action: fd.id ? 'update' : 'create', fd: fd, newAccount: newAccount });
        if (ok) setEditing(null);
        return ok;
      }
    });
  }

  function handleAccountSave(account, originalName){
    const renamed = originalName && originalName !== account.name;
    const cascade = renamed
      ? data.fds.filter(function(f){ return f.account === originalName; }).length
      : 0;
    setConfirm({
      title: originalName ? 'Save changes to this account?' : 'Add this account?',
      lines: [
        account.name,
        [(account.person || '—'), (account.bank || '—')].join(' · '),
        cascade > 0 ? 'Renaming also updates ' + cascade + ' FD' + (cascade === 1 ? '' : 's') + ' under “' + originalName + '”.' : ''
      ],
      confirmLabel: originalName ? 'Yes, save changes' : 'Yes, add account',
      danger: false,
      run: async function(){
        const ok = await mutate(originalName
          ? { action:'accountUpdate', account: account, originalName: originalName }
          : { action:'accountCreate', account: account });
        if (ok) setAccEditing(null);
        return ok;
      }
    });
  }

  function accountFdCount(name){
    return data.fds.filter(function(f){ return String(f.account).toLowerCase() === String(name).toLowerCase(); }).length;
  }

  function handleAccountDelete(acc){
    /* an account holding FDs can't be deleted until they're moved — so the
       Delete button opens the transfer step first (the user deletes after) */
    if (accountFdCount(acc.name) > 0){
      setTransferring({ account: acc, forceOut: true });
      return;
    }
    setConfirm({
      title: 'Delete this account?',
      lines: [
        acc.name,
        [(acc.person || '—'), (acc.bank || '—')].join(' · '),
        LIVE ? 'This removes the row from the Accounts tab of the sheet.' : 'Demo: it comes back on refresh.'
      ],
      confirmLabel: 'Yes, delete account',
      danger: true,
      run: function(){ return mutate({ action:'accountDelete', name: acc.name }); }
    });
  }

  /* TransferDialog → confirm → move. Confirmation always precedes the shift. */
  function handleTransferSubmit(ids, from, to, moving){
    setTransferring(null);
    setConfirm({
      title: 'Move ' + moving + ' FD' + (moving === 1 ? '' : 's') + '?',
      lines: [
        'From “' + from + '” to “' + to + '”.',
        'Each selected FD\'s account is reassigned; nothing else changes.',
        LIVE ? 'Updates those rows in the Google Sheet.' : 'Demo: resets on refresh.'
      ],
      confirmLabel: 'Yes, move ' + (moving === 1 ? 'it' : 'them'),
      danger: false,
      run: function(){ return mutate({ action:'transferFds', ids: ids, from: from, to: to }); }
    });
  }

  function handleDelete(fd){
    setConfirm({
      title: 'Delete this FD?',
      lines: fdSummary(fd).concat([
        LIVE ? 'This removes the row from the Google Sheet.' : 'Demo: it comes back on refresh.'
      ]),
      confirmLabel: 'Yes, delete FD',
      danger: true,
      run: function(){ return mutate({ action:'delete', id: fd.id }); }
    });
  }

  /* "Mark as done" — one tap, no confirmation (fully reversible via History):
     a matured FD is archived to Inactive and drops off the dashboard. Uses a
     targeted markDone action (not a full update) so no other cell is touched. */
  function handleMarkDone(fd){
    setConfirm({
      title: 'Mark this FD as done?',
      lines: fdSummary(fd).concat(['Moves it to Inactive, off the dashboard. You can undo it any time from History.']),
      confirmLabel: 'Yes, mark as done',
      danger: false,
      run: function(){
        return mutate({ action: 'markDone', id: fd.id },
          { saved: 'Marked FD' + (fd.accountNumber ? ' ' + fd.accountNumber : '') + ' as done — moved to Inactive' });
      }
    });
  }

  function handleUndo(entry){
    setConfirm({
      title: 'Undo this change?',
      lines: [
        entry.summary,
        fmtStamp(entry.time) + (entry.device ? ' · ' + entry.device : ''),
        'Restores the previous values.'
      ],
      confirmLabel: 'Yes, undo',
      danger: false,
      run: async function(){
        const ok = await mutate({ action:'undo', logId: entry.id }, { saved: 'Undone: ' + entry.summary, flash: flashFromTarget(entry.target) });
        /* refused: the list on screen was out of date — re-pull it quietly, so
           an Undo link that can no longer work disappears */
        if (!ok) refresh({ quiet: true });
        return ok;
      }
    });
  }

  const VERBS = { create:'Added', update:'Edited', delete:'Deleted', undo:'Restored',
    accountCreate:'Added', accountUpdate:'Edited', accountDelete:'Deleted' };

  /* Background reload from the sheet; never blocks the UI, announces itself. */
  /* opts.force: run even if a refresh is already in flight (entering History,
   * whose month may differ from the one that refresh asked for).
   * opts.wantInitiator: this refresh is the one fetching a newly wanted month;
   * if it fails, History falls back to the month on screen.
   * opts.quiet: no toast either way (a follow-up re-pull after a refused Undo).
   * The top-bar button passes its click event here, which is none of these. */
  async function refresh(opts){
    const force = !!(opts && opts.force === true);
    const quiet = !!(opts && opts.quiet === true);
    if (!LIVE || (refreshing && !force)) return;
    setRefreshing(true);
    const asked = wantRef.current;
    try {
      const j = await call(withLogMonth({ action:'list', password: auth.password }));
      applyServer(j, asked);
      /* never over a save: not while one runs (its "Saving…" stays up), and
         not in place of the toast confirming one */
      if (!quiet && inflightRef.current.saves === 0){
        setToast(function(t){ return t && t.save ? t : { msg: 'Refreshed', kind: 'success' }; });
      }
    } catch (e) {
      if (opts && opts.wantInitiator === true) rollBackWant(asked);
      if (!quiet) notifyError('Refresh failed: ' + e.message);
    } finally {
      setRefreshing(false);
    }
  }

  /* Switching views re-pulls the sheet so the numbers are never stale. */
  /* History: show another month, fetched quietly (no "Refreshed" toast). */
  async function pickLogMonth(m){
    const asked = wantMonth(m);
    if (!LIVE){ setData(function(prev){ return prev ? demoLogView(prev) : prev; }); markLoaded(asked.seq); return; }
    try {
      const j = await call(withLogMonth({ action:'list', password: auth.password }));
      applyServer(j, asked);
    } catch (e) {
      rollBackWant(asked);
      notifyError('Could not load that month: ' + e.message);
    }
  }

  /* Switching tabs re-pulls the sheet only when what's shown is over a minute
   * old — every save already brings back the fresh sheet, so this just skips
   * reloads that couldn't show anything new. */
  const STALE_MS = 60 * 1000;
  function stale(){ return Date.now() - lastSyncRef.current >= STALE_MS; }

  function switchTab(t){
    if (t !== tab){
      if (t === 'history'){
        const reset = !wantsThisMonth(); /* History always opens on this month */
        if (reset){
          const w = wantMonth(null);
          if (!LIVE){ setData(function(prev){ return prev ? demoLogView(prev) : prev; }); markLoaded(w.seq); }
        }
        if (LIVE && (reset || stale())) refresh({ force: reset, wantInitiator: reset });
      } else if (LIVE && stale()){
        refresh();
      }
    }
    setTab(t);
    FDPrefs.setTab(t); /* survive a page refresh */
  }

  if (boot){
    return (
      <div className="login-wrap">
        <div className="login-card">
          <div className="login-mark" aria-hidden="true">₹</div>
          <h1>Family FD Tracker</h1>
          <p className="login-sub">Opening the ledger with your saved sign-in…</p>
        </div>
      </div>
    );
  }

  if (!auth) return <Login onLogin={login} busy={busy} error={loginErr} />;

  const accountsByName = {};
  data.accounts.forEach(function(a){ accountsByName[a.name] = a; });
  const joined = data.fds.map(function(f){
    const a = accountsByName[f.account];
    return Object.assign({}, f, { person: a ? a.person : '—', bank: a ? a.bank : '—' });
  });
  const shownCount = joined.filter(isOnDashboard).length;

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark" aria-hidden="true">₹</span>
          <div>
            <div className="brand-title">Family FD Tracker</div>
            <div className="brand-sub">{joined.length} deposits · {shownCount} on dashboard</div>
          </div>
        </div>
        <span className={'mode-badge ' + (LIVE ? 'live' : 'demo')}>{LIVE ? 'LIVE' : 'DEMO'}</span>
        <nav className="tabs" aria-label="Views">
          <button type="button" className={'tab-btn' + (tab === 'dashboard' ? ' on' : '')}
                  onClick={function(){ switchTab('dashboard'); }}>Dashboard</button>
          <button type="button" className={'tab-btn' + (tab === 'all' ? ' on' : '')}
                  onClick={function(){ switchTab('all'); setAllKey(function(k){ return k + 1; }); }}>All FDs</button>
          <button type="button" className={'tab-btn' + (tab === 'accounts' ? ' on' : '')}
                  onClick={function(){ switchTab('accounts'); }}>Accounts</button>
          <button type="button" className={'tab-btn' + (tab === 'history' ? ' on' : '')}
                  onClick={function(){ if (tab === 'history'){ if (!wantsThisMonth()) pickLogMonth(null); } else switchTab('history'); }}>History</button>
          {LIVE && (
            <button type="button" className={'tab-btn' + (tab === 'reminders' ? ' on' : '')}
                    onClick={function(){ switchTab('reminders'); }}>Reminders</button>
          )}
        </nav>
        <div className="topbar-actions">
          <button type="button" className="btn theme-btn" onClick={toggleTheme}
                  title="Switch between light and dark">
            {theme === 'dark' ? 'Light mode' : 'Dark mode'}
          </button>
          {LIVE && (
            <button type="button" className="btn" title="Reload from the Google Sheet"
                    onClick={refresh} disabled={refreshing} aria-busy={refreshing}>
              <span className={'spin-icon' + (refreshing ? ' spinning' : '')} aria-hidden="true">↻</span> Refresh
            </button>
          )}
          <button type="button" className="btn primary" onClick={function(){ setEditing('new'); }}>＋ Add FD</button>
          <button type="button" className="btn" onClick={logout} title="Forget the saved sign-in on this device">
            Sign out
          </button>
        </div>
      </header>

      {!LIVE && (
        <div className="demo-banner">
          <b>Demo mode.</b> This is sample data — changes are not saved and reset on refresh.
          Paste your Apps Script URL into <code>APPS_SCRIPT_URL</code> in <code>index.html</code> to go live (see SETUP.md).
        </div>
      )}

      <main>
        {tab === 'dashboard' && <Dashboard fds={joined} flash={flash} onEdit={function(f){ setEditing(f); }} onMarkDone={handleMarkDone} />}
        {tab === 'all' && <AllFDs key={allKey} fds={joined} accounts={data.accounts} flash={flash} onEdit={function(f){ setEditing(f); }} onDelete={handleDelete} onMarkDone={handleMarkDone} />}
        {tab === 'accounts' && <AccountsView accounts={data.accounts} fds={joined} flash={flash}
                onAdd={function(){ setAccEditing('new'); }}
                onEdit={function(a){ setAccEditing(a); }}
                onTransfer={function(a){ setTransferring({ account: a, forceOut: false }); }}
                onDelete={handleAccountDelete} />}
        {tab === 'history' && <RecentChanges log={data.log || []} logMonth={data.logMonth} logMonths={data.logMonths}
                                            loading={want.seq > loadedSeq} wantMonth={want.month}
                                            onPickMonth={pickLogMonth} onUndo={handleUndo} onView={function(e){ setViewing(e); }} />}
        {tab === 'reminders' && <RemindersView pushSubs={data.pushSubs || []} pushState={pushState}
                thisEndpoint={thisEndpoint} busy={pushState === 'busy'}
                onEnable={enableReminders} onDisable={disableReminders} onUnsubscribe={unsubscribeDevice} />}
      </main>

      {editing !== null && (
        <FdForm initial={editing === 'new' ? null : editing}
                accounts={data.accounts}
                busy={busy}
                onCancel={function(){ setEditing(null); }}
                onSave={handleSave} />
      )}

      {accEditing !== null && (
        <AccountForm initial={accEditing === 'new' ? null : accEditing}
                accounts={data.accounts}
                fdCount={accEditing === 'new' ? 0 : joined.filter(function(f){ return f.account === accEditing.name; }).length}
                busy={busy}
                onCancel={function(){ setAccEditing(null); }}
                onSave={handleAccountSave} />
      )}

      {transferring !== null && (
        <TransferDialog account={transferring.account} accounts={data.accounts} fds={joined}
                forceOut={transferring.forceOut} busy={busy}
                onCancel={function(){ setTransferring(null); }}
                onSubmit={handleTransferSubmit} />
      )}

      {confirm !== null && (
        <ConfirmDialog title={confirm.title} lines={confirm.lines}
                confirmLabel={confirm.confirmLabel} danger={confirm.danger}
                busy={busy} error={confirm.error}
                onConfirm={runConfirm}
                onCancel={function(){ setConfirm(null); }} />
      )}

      {viewing !== null && (function(){
        const s = logSnapshot(viewing);
        if (!s) return null;
        const snap = s.snap, isFd = s.isFd;
        const verb = VERBS[viewing.action] || 'Recorded';
        const note = 'Read-only — exactly what was recorded · ' +
          fmtStamp(viewing.time) + (viewing.device ? ' · ' + viewing.device : '');
        const close = function(){ setViewing(null); };
        return isFd ? (
          <FdForm initial={snap} accounts={data.accounts} busy={false}
                  readOnly={true}
                  title={verb + ' FD' + (snap.id ? ' #' + snap.id : '')}
                  headNote={note}
                  onCancel={close} onSave={function(){}} />
        ) : (
          <AccountForm initial={snap} accounts={data.accounts} fdCount={0} busy={false}
                  readOnly={true}
                  title={verb + ' account'}
                  headNote={note}
                  onCancel={close} onSave={function(){}} />
        );
      })()}

      {inflight.all > 0 && <div className="sync-bar" aria-hidden="true"><div className="sync-bar-fill" /></div>}
      <div className="toast-stack">
        {toast && (
          <div className={'toast ' + toast.kind} role={toast.kind === 'error' ? 'alert' : 'status'}>
            <span>{toast.kind === 'success' ? '✓ ' : ''}{toast.msg}</span>
            <button type="button" aria-label="Dismiss" onClick={function(){ setToast(null); }}>✕</button>
          </div>
        )}
        {inflight.all > 0 && (
          <div className="toast pending" role="status">
            <span className="sync-dot" aria-hidden="true" />{inflight.saves > 0 ? 'Saving…' : 'Loading…'}
          </div>
        )}
      </div>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
