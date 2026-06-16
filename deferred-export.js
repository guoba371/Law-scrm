"use strict";

function createDeferredRunner(run, options = {}) {
  const schedule = options.schedule || (callback => setTimeout(callback, 0));
  const cancel = options.cancel || (handle => clearTimeout(handle));

  let pendingHandle = null;
  let pendingArgs = [];

  return function deferredRunner(...args) {
    pendingArgs = args;
    if (pendingHandle !== null) {
      cancel(pendingHandle);
    }
    pendingHandle = schedule(() => {
      pendingHandle = null;
      run(...pendingArgs);
    });
  };
}

module.exports = { createDeferredRunner };
