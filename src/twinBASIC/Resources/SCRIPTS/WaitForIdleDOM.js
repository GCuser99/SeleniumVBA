async function waitForIdleDOM(root, idleMs = 50, timeout = 30000) {
try {
  // Normalize root argument
  if (root === 'document') {
    root = document;
  }
  if (!root || !(root instanceof Node)) {
    return 'Error: invalid root node';
  }
  let lastMutation = Date.now();
  const observer = new MutationObserver(() => {
    lastMutation = Date.now();
  });
  observer.observe(root, {
    subtree: true,
    childList: true,
    attributes: true,
    characterData: true
  });
  const checkInterval = Math.min(25, idleMs);
  return await new Promise(resolve => {
    const interval = setInterval(() => {
      if (Date.now() - lastMutation >= idleMs) {
        cleanup();
        resolve('done');
      }
    }, checkInterval);
    const safety = setTimeout(() => {
      cleanup();
      resolve('Error: timeout waiting for DOM idle');
    }, timeout);
    function cleanup() {
      clearInterval(interval);
      clearTimeout(safety);
      observer.disconnect();
    }
  });
} catch (err) {
  return `Error: ${err && err.message ? err.message : String(err)}`;
}
}
return waitForIdleDOM(arguments[0], arguments[1], arguments[2]);