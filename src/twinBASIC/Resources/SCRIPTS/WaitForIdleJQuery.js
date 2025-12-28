var callback = arguments[arguments.length - 1]; async function waitForIdleJQuery(timeout = 10000, poll = 50) {
if (!window.jQuery || typeof jQuery.active !== "number") {
  return 'jQuery not present or jQuery.active unavailable';
}
const start = Date.now();
while (jQuery.active > 0) {
  if (Date.now() - start > timeout) {
    return 'Error: timeout waiting for jQuery idle';
  }
  await new Promise(r => setTimeout(r, poll));
}
return 'done';
}
return waitForIdleJQuery(arguments[0], arguments[1]).then(callback);