const expected = ['static', 'test', 'browser', 'fidelity', 'ooxml-oracle'];
const value: unknown = JSON.parse(process.env.CI_JOB_RESULTS ?? 'null');
if (typeof value !== 'object' || value === null || Array.isArray(value)) {
  throw new Error('Missing required CI job results');
}
const results = new Map(Object.entries(value));
if (results.size !== expected.length) throw new Error('Unexpected required job inventory');
for (const name of expected) {
  const job: unknown = results.get(name);
  if (typeof job !== 'object' || job === null || !('result' in job) || job.result !== 'success') {
    throw new Error(`Required CI job did not succeed: ${name}`);
  }
}
console.log('All required CI jobs succeeded (including every matrix member)');
