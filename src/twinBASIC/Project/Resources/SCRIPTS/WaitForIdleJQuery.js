// Args (all optional): [0] timeout ms, [1] poll ms, [2] settle ms. The last argument is always the WebDriver callback.
var callback = arguments[arguments.length - 1];
var userArgs = Array.prototype.slice.call(arguments, 0, arguments.length - 1);

async function waitForIdleJQuery(timeout = 10000, poll = 50, settle = 100) {
  const start = performance.now();
  let idleSince = null;
  while (true) {
    const jq = window.jQuery; // re-read each poll; page may load/replace jQuery
    if (!jq || typeof jq.active !== "number") {
      return 'jQuery not present or jQuery.active unavailable';
    }
    const now = performance.now();
    if (jq.active === 0) {
      if (idleSince === null) idleSince = now;
      if (now - idleSince >= settle) return 'done';
    } else {
      idleSince = null; // a new request started during the settle window
    }
    if (now - start > timeout) {
      return jq.active === 0
        ? 'done' // idle but settle window not yet complete: don't report a false timeout
        : 'Error: timeout waiting for jQuery idle (jQuery.active = ' + jq.active + ')';
    }
    await new Promise(r => setTimeout(r, poll));
  }
}

waitForIdleJQuery.apply(null, userArgs).then(
  callback,
  e => callback('Error: exception in WaitForIdleJQuery: ' + (e && e.message ? e.message : e))
);