import { spawn } from 'node:child_process';
import { mkdir, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { requireSuites, testFiles, type SuiteResult } from './suite-results.ts';

// The reference is a pinned submodule, and JSZip is locked by the root workspace.
// Prove it can load before Vitest can turn its availability guard into a skip.
createRequire(import.meta.url)('../../references/PptxGenJS/dist/pptxgen.cjs.js');
const expected = [
  ...(await testFiles(resolve('src'), true)),
  ...(await testFiles(resolve('test'), true)),
];
const report = resolve('.probe/ci/core.json');
await mkdir(resolve('.probe/ci'), { recursive: true });
// The opt-in native Windows font scan competes with other suite workers for
// hosted-runner I/O and memory. Isolate it without changing tests or deadlines.
const workers = process.platform === 'win32' ? 1 : 2;
console.log(`Running the complete core suite with ${workers} worker(s)`);
const exitCode = await new Promise<number | null>((done, reject) => {
  const child = spawn(
    process.execPath,
    [
      fileURLToPath(new URL('./vitest.mjs', import.meta.resolve('vitest/package.json'))),
      'run',
      `--maxWorkers=${workers}`,
      '--reporter=default',
      '--reporter=json',
      `--outputFile.json=${report}`,
    ],
    {
      stdio: 'inherit',
      env: { ...process.env, PERF: '1', GENERATE_SAMPLES: '1' },
    },
  );
  child.on('error', reject);
  child.on('close', done);
});
if (exitCode !== 0) throw new Error(`Vitest failed: ${exitCode}`);

function record(value: unknown): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Error('Invalid test report');
  return Object.fromEntries(Object.entries(value));
}
const data: unknown = JSON.parse(await readFile(report, 'utf8'));
const testResults = record(data).testResults;
if (!Array.isArray(testResults)) throw new Error('Missing Vitest testResults');
const results: SuiteResult[] = testResults.map((value: unknown) => {
  const suite = record(value);
  if (typeof suite.name !== 'string' || !Array.isArray(suite.assertionResults))
    throw new Error('Invalid suite report');
  const result: SuiteResult = { file: suite.name, passed: 0, failed: 0, skipped: 0, todo: 0 };
  for (const assertion of suite.assertionResults) {
    const status = record(assertion).status;
    if (status === 'passed') result.passed++;
    else if (status === 'failed') result.failed++;
    else if (status === 'pending' || status === 'skipped') result.skipped++;
    else if (status === 'todo') result.todo++;
    else throw new Error(`Unknown assertion status: ${String(status)}`);
  }
  return result;
});
requireSuites(expected, results);
console.log(
  `Verified ${expected.length} executed Vitest files, including samples, performance and reference corpus`,
);
