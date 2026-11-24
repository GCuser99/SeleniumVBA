function waitForIdleNetwork (idleTimeout=1500, maxTimeout=30000){
return new Promise((resolve, reject) => {
  let req = 0, timer = null, settled = false;
  const checkInterval = 100;
  const of = window.fetch;
  const oo = XMLHttpRequest.prototype.open;
  const os = XMLHttpRequest.prototype.send;
  function resetTimer() { if (timer) clearTimeout(timer); }
  function checkIdle() {
    resetTimer();
    if (req === 0) {
      timer = setTimeout(() => {
        window.fetch = of;
        XMLHttpRequest.prototype.open = oo;
        XMLHttpRequest.prototype.send = os;
        if (!settled) { settled = true; clearTimeout(safety); resolve(); }
      }, idleTimeout);
    }
  }
  window.fetch = function(u, o) {
    req++; resetTimer();
    return of(u, o).finally(() => { req--; checkIdle(); });
  };
  XMLHttpRequest.prototype.open = function(...a) { return oo.apply(this, a); };
  XMLHttpRequest.prototype.send = function(...a) {
    req++; resetTimer();
    this.addEventListener('loadend', () => { req--; checkIdle(); });
    return os.apply(this, a);
  };
  setTimeout(checkIdle, checkInterval);
  const safety = setTimeout(() => {
    if (!settled) { settled = true; reject(new Error("networkIdle timeout")); }
  }, maxTimeout);
});
}
return waitForIdleNetwork(arguments[0], arguments[1]);