// Args (optional): [0] array of URL substrings to ignore (long-polling, analytics, websockets-over-HTTP, etc.)
if (!window.__niw) { (function (ignore) {
  ignore = Array.isArray(ignore) ? ignore : [];
  var s = { req: 0, lastAct: Date.now(), onchange: null };
  function skip(url) {
    url = String(url || '');
    return ignore.some(function (p) { return url.indexOf(p) !== -1; });
  }
  // Every request start/end and every DOM mutation counts as "activity"
  function bump(d) { s.req += d; s.lastAct = Date.now(); if (s.onchange) s.onchange(); }
  function tracked(p) { // decrement exactly once when promise p settles
    return p.finally(function () { bump(-1); });
  }

  var of = window.fetch;
  if (of) {
    window.fetch = function (input) {
      if (skip(input && input.url ? input.url : input)) return of.apply(this, arguments);
      bump(1);
      try { return tracked(of.apply(this, arguments)); } catch (e) { bump(-1); throw e; }
    };
    // fetch() resolves when headers arrive; the body may still be downloading.
    // Count body reads as pending work so the quiet window starts after the body is in.
    ['json', 'text', 'blob', 'arrayBuffer', 'formData'].forEach(function (m) {
      var orig = window.Response && Response.prototype[m];
      if (!orig) return;
      Response.prototype[m] = function () {
        bump(1);
        try { return tracked(orig.apply(this, arguments)); } catch (e) { bump(-1); throw e; }
      };
    });
  }

  var oo = XMLHttpRequest.prototype.open, os = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (method, url) {
    this.__niwSkip = skip(url);
    return oo.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function () {
    if (this.__niwSkip) return os.apply(this, arguments);
    var ended = false;
    function end() { if (!ended) { ended = true; bump(-1); } }
    bump(1);
    this.addEventListener('loadend', end);
    try { return os.apply(this, arguments); } catch (e) { end(); throw e; } // e.g. send() before open()
  };

  new MutationObserver(function () {
    s.lastAct = Date.now();
    if (s.onchange) s.onchange();
  }).observe(document, { childList: true, subtree: true, attributes: true, characterData: true });

  window.__niw = s;
})(arguments[0]); }