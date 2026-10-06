export function collectRejectedReasons<T>(settled: PromiseSettledResult<T>[]): unknown[] {
  return settled
    .filter(function(result): result is PromiseRejectedResult {
      return result.status === 'rejected';
    })
    .map(function(result) {
      return result.reason;
    });
}

export function collectCreatedIds<T extends { data: { Id: number } }>(
  settled: PromiseSettledResult<T>[]
): { ids: number[]; errors: unknown[] } {
  const ids: number[] = [];
  const errors: unknown[] = [];

  settled.forEach(function(result) {
    if (result.status === 'fulfilled') {
      ids.push(result.value.data.Id);
    } else {
      errors.push(result.reason);
    }
  });

  return { ids, errors };
}

export function createBatchError(action: 'create' | 'update' | 'delete', failed: number, total: number): Error {
  return new Error(`Batch ${action} failed: ${failed} of ${total} items failed`);
}
