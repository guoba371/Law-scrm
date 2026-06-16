const test = require("node:test");
const assert = require("node:assert/strict");

const { createDeferredRunner } = require("./deferred-export.js");

test("coalesces rapid calls into one deferred run", () => {
  const queued = new Map();
  const canceled = [];
  let runs = 0;
  let nextId = 0;

  const runner = createDeferredRunner(
    value => {
      runs += 1;
      return value;
    },
    {
      schedule(callback) {
        nextId += 1;
        queued.set(nextId, callback);
        return nextId;
      },
      cancel(handle) {
        canceled.push(handle);
        queued.delete(handle);
      }
    }
  );

  runner("a");
  runner("b");
  runner("c");

  assert.equal(runs, 0);
  assert.deepEqual(canceled, [1, 2]);

  queued.forEach(callback => callback());

  assert.equal(runs, 1);
});

test("uses the latest arguments for the deferred run", () => {
  const queued = [];
  const seen = [];

  const runner = createDeferredRunner(
    (...args) => {
      seen.push(args);
    },
    {
      schedule(callback) {
        queued.push(callback);
        return queued.length;
      },
      cancel() {}
    }
  );

  runner("first", 1);
  runner("second", 2);

  queued[queued.length - 1]();

  assert.deepEqual(seen, [["second", 2]]);
});
