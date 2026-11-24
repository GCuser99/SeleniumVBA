// Args (optional): [0] idle ms, [1] max ms. The last argument is always the WebDriver callback.
// Returns 'ok', 'timeout', or 'not-armed' (hooks not installed in this document, e.g. the page navigated).
var cb = arguments[arguments.length - 1];
var a = Array.prototype.slice.call(arguments, 0, arguments.length - 1);
var idleTimeout = a[0] != null ? a[0] : 500;
var maxTimeout = a[1] != null ? a[1] : 30000;

var s = window.__niw;
if (!s) { cb('not-armed'); return; }

var tm = null, settled = false;
var safety = setTimeout(function () { done('timeout'); }, maxTimeout);

function done(v) {
  if (settled) return;
  settled = true;
  clearTimeout(tm);
  clearTimeout(safety);
  if (s.onchange === schedule) s.onchange = null; // don't leave a stale handler running on every mutation
  cb(v);
}

function schedule() {
  clearTimeout(tm);
  if (s.req > 0) return; // the next request end will call us again
  var remaining = Math.max(0, idleTimeout - (Date.now() - s.lastAct));
  tm = setTimeout(function () {
    if (s.req === 0 && Date.now() - s.lastAct >= idleTimeout) done('ok');
    else schedule();
  }, remaining);
}

s.onchange = schedule;
schedule();