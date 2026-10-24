import { spawn } from 'node:child_process';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { testFiles, type SuiteResult } from './suite-results.ts';
import { requireVitestPartitions, type VitestPartition } from './vitest-partitions.ts';

// The reference is a pinned submodule, and JSZip is locked by the root workspace.
// Prove it can load before Vitest can turn its availability guard into a skip.
createRequire(import.meta.url)('../../references/PptxGenJS/dist/pptxgen.cjs.js');
const expected = [
  ...(await testFiles(resolve('src'), true)),
  ...(await testFiles(resolve('test'), true)),
];
const rasterFilter = 'test/preview-node-raster.test.ts';
const rasterFile = resolve(rasterFilter);
if (!expected.includes(rasterFile)) throw new Error('Missing required Node raster suite');
const report = resolve('.probe/ci/core.json');
await mkdir(resolve('.probe/ci'), { recursive: true });
await rm(report, { force: true });

// The native Windows font database is constructed synchronously. Run its whole
// suite in a fresh process before the general worker accumulates native state.
// Every source file belongs to exactly one mandatory partition, with the same
// tests, PERF/sample flags and Vitest deadlines as the unpartitioned run.
const plans =
  process.platform === 'win32'
    ? [
        { name: 'raster', files: [rasterFile], args: [rasterFilter] },
        {
          name: 'suite',
          files: expected.filter((file) => file !== rasterFile),
          args: ['--exclude', rasterFilter],
        },
      ]
    : [{ name: 'suite', files: expected, args: [] }];
const workers = process.platform === 'win32' ? 1 : 2;
const vitest = fileURLToPath(new URL('./vitest.mjs', import.meta.resolve('vitest/package.json')));
const run = (args: string[]): Promise<number | null> =>
  new Promise((done, reject) => {
    const child = spawn(process.execPath, args, {
      stdio: 'inherit',
      env: { ...process.env, PERF: '1', GENERATE_SAMPLES: '1' },
    });
    child.on('error', reject);
    child.on('close', done);
  });
const executions = [];
for (const plan of plans) {
  const path = resolve(`.probe/ci/core-${plan.name}.json`);
  await rm(path, { force: true });
  console.log(`Running ${plan.name}: ${plan.files.length} required file(s), ${workers} worker(s)`);
  const exitCode = await run([
    vitest,
    'run',
    ...plan.args,
    `--maxWorkers=${workers}`,
    '--reporter=default',
    '--reporter=json',
    `--outputFile.json=${path}`,
  ]);
  executions.push({ ...plan, path, exitCode });
}

// Diagnostics run after the actual test partitions, so they cannot warm the
// font scan that the 5-second raster test is required to complete on its own.
if (process.platform === 'win32') {
  const exitCode = await run(['packages/preview/scripts/diagnose-raster-fonts.ts']);
  if (exitCode !== 0) throw new Error(`Native raster diagnostics failed: ${exitCode}`);
}

function record(value: unknown): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Error('Invalid test report');
  return Object.fromEntries(Object.entries(value));
}
const partitions: VitestPartition[] = [];
const testResults: Record<string, unknown>[] = [];
const rawReports: Record<string, unknown>[] = [];
for (const execution of executions) {
  let parsed: unknown;
  try {
    parsed = JSON.parse(await readFile(execution.path, 'utf8'));
  } catch (cause) {
    throw new Error(`Missing or invalid execution report for ${execution.name}`, { cause });
  }
  const data = record(parsed);
  rawReports.push(data);
  if (!Array.isArray(data.testResults)) throw new Error('Missing Vitest testResults');
  const results: SuiteResult[] = data.testResults.map((value: unknown) => {
    const suite = record(value);
    if (typeof suite.name !== 'string' || !Array.isArray(suite.assertionResults))
      throw new Error('Invalid suite report');
    testResults.push(suite);
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
  partitions.push({ ...execution, results });
}
const results = requireVitestPartitions(expected, partitions);
const countFields = [
  'numTotalTestSuites',
  'numPassedTestSuites',
  'numFailedTestSuites',
  'numPendingTestSuites',
  'numTotalTests',
  'numPassedTests',
  'numFailedTests',
  'numPendingTests',
  'numTodoTests',
];
const totals = Object.fromEntries(
  countFields.map((key) => [
    key,
    rawReports.reduce((sum, data) => {
      const value = data[key];
      if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0)
        throw new Error(`Invalid Vitest report count: ${key}`);
      return sum + value;
    }, 0),
  ]),
);
const passedTests = results.reduce((sum, suite) => sum + suite.passed, 0);
if (
  totals.numTotalTests !== passedTests ||
  totals.numPassedTests !== passedTests ||
  totals.numFailedTests !== 0 ||
  totals.numPendingTests !== 0 ||
  totals.numTodoTests !== 0
) {
  throw new Error('Vitest report totals disagree with the complete assertion inventory');
}
await writeFile(report, JSON.stringify({ ...totals, success: true, testResults }, null, 2));
console.log(
  `Verified ${expected.length} executed Vitest files and ${passedTests} tests, including samples, performance and reference corpus`,
);
