/** Execute a runner with immediately registered, awaited LIFO cleanup.
 * defer returns the same idempotent async disposer that teardown will invoke.
 * With no cleanup errors, the original rejection is rethrown unchanged.
 * Cleanup errors reject with AggregateError: errors are [primary, ...cleanup]
 * and cause is the original primary when run failed; otherwise errors are the
 * cleanup failures. This keeps reporting failures observable without preventing
 * the remaining resources from closing (including frozen/primitive primaries).
 */
async function withBrowserCleanup(run) {
  const disposers = [];
  const defer = dispose => {
    let pending;
    const once = () => {
      if (!pending) pending = Promise.resolve().then(dispose);
      return pending;
    };
    disposers.push(once);
    return once;
  };
  let value, primary, failed = false;
  try { value = await run(defer); }
  catch (error) { primary = error; failed = true; }
  const errors = [];
  for (const dispose of disposers.reverse()) {
    try { await dispose(); } catch (error) { errors.push(error); }
  }
  if (errors.length) throw new AggregateError(failed ? [primary, ...errors] : errors, 'Browser runner cleanup failed', failed ? { cause: primary } : undefined);
  if (failed) throw primary;
  return value;
}
module.exports = { withBrowserCleanup };
