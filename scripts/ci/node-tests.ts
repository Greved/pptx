import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { run } from 'node:test';
import { spec } from 'node:test/reporters';
import { requireSuites, testFiles, type SuiteResult } from './suite-results.ts';

const suites = {
  'dev-unit': 'packages/dev/test',
  'editor-unit': 'packages/editor/test',
  'dev-browser': 'packages/dev/test/browser',
  'editor-browser': 'packages/editor/test/browser',
};
const suite = process.argv[2];
if (
  suite !== 'dev-unit' &&
  suite !== 'editor-unit' &&
  suite !== 'dev-browser' &&
  suite !== 'editor-browser'
) {
  throw new Error('Expected dev-unit, editor-unit, dev-browser or editor-browser');
}
const directory = resolve(suites[suite]);
const allFiles = await testFiles(directory);
if (suite === 'editor-browser') {
  for (const name of ['embed.test.mjs', 'render-performance.test.ts', 'svg-paint.test.ts']) {
    if (!allFiles.includes(resolve(directory, name))) {
      throw new Error(`Missing required embedding/render performance suite: ${name}`);
    }
  }
}
// Keep eight complete, disjoint shards, serial execution and the existing timeout.
const shardArg = process.argv[3];
let shard: { index: number; total: number } | undefined;
if (shardArg !== undefined) {
  if (suite !== 'dev-browser') throw new Error('Only dev-browser is sharded');
  const match = /^(\d+)\/(\d+)$/.exec(shardArg);
  if (!match?.[1] || !match[2]) throw new Error('Invalid shard');
  const index = Number(match[1]);
  const total = Number(match[2]);
  if (!Number.isSafeInteger(index) || total !== 8 || index < 1 || total < index) {
    throw new Error('Expected one of the eight dev-browser shards');
  }
  shard = { index, total };
}
// Explicit partitioning also makes the expected files independent of runner events.
const start = shard ? Math.floor((allFiles.length * (shard.index - 1)) / shard.total) : 0;
const end = shard ? Math.floor((allFiles.length * shard.index) / shard.total) : allFiles.length;
const files = allFiles.slice(start, end);
if (files.length === 0) throw new Error('Empty test shard');
const results: SuiteResult[] = [];
let successful = false;
const stream = run({
  cwd: resolve(directory, suite.endsWith('browser') ? '../..' : '..'),
  files,
  concurrency: 1,
  ...(suite.endsWith('browser') ? { timeout: 180_000 } : {}),
});
stream.on('test:summary', (summary) => {
  if (summary.file === undefined) successful = summary.success;
  else {
    results.push({
      file: summary.file,
      passed: summary.counts.passed,
      failed:
        summary.counts.tests - summary.counts.passed - summary.counts.skipped - summary.counts.todo,
      skipped: summary.counts.skipped,
      todo: summary.counts.todo,
    });
  }
});
await new Promise<void>((done, reject) => {
  const output = stream.compose(spec);
  output.on('error', reject);
  output.on('end', done);
  output.pipe(process.stdout, { end: false });
});
const outputDirectory = resolve('.probe/ci');
await mkdir(outputDirectory, { recursive: true });
await writeFile(
  resolve(outputDirectory, `${suite}${shard ? `-${shard.index}` : ''}.json`),
  JSON.stringify(results, null, 2),
);
requireSuites(files, results);
if (!successful) throw new Error('Test runner did not report successful completion');
console.log(`Verified ${files.length} executed test files for ${suite}`);
