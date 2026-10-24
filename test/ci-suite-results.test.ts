import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { describe, expect, it } from 'vitest';
import { requireSuites, testFiles, type SuiteResult } from '../scripts/ci/suite-results.ts';
import { requireVitestPartitions, type VitestPartition } from '../scripts/ci/vitest-partitions.ts';

const file = resolve('example.test.ts');
const passed: SuiteResult = { file, passed: 1, failed: 0, skipped: 0, todo: 0 };
const execute = promisify(execFile);

describe('CI execution evidence', () => {
  it('requires every process partition and its own exact source inventory', () => {
    const rasterFile = resolve('test/preview-node-raster.test.ts');
    const raster: SuiteResult = { ...passed, file: rasterFile, passed: 10 };
    const expected = [file, rasterFile];
    const core: VitestPartition = { name: 'core', files: [file], results: [passed], exitCode: 0 };
    const native: VitestPartition = {
      name: 'raster',
      files: [rasterFile],
      results: [raster],
      exitCode: 0,
    };
    const partitions = [core, native];
    expect(requireVitestPartitions(expected, partitions)).toEqual([passed, raster]);
    for (const replacement of [
      { ...native, results: undefined },
      { ...native, exitCode: 1 },
      { ...native, exitCode: null },
      { ...native, results: [] },
      { ...native, results: [{ ...raster, failed: 1 }] },
      { ...native, results: [{ ...raster, skipped: 1 }] },
      { ...native, results: [{ ...raster, todo: 1 }] },
      { ...native, results: [passed, raster] },
    ]) {
      expect(() => requireVitestPartitions(expected, [core, replacement])).toThrow();
    }
    expect(() => requireVitestPartitions(expected, partitions.slice(0, 1))).toThrow(
      'Missing test result',
    );
    expect(() =>
      requireVitestPartitions(expected, [
        { name: 'all', files: expected, results: [passed, raster], exitCode: 0 },
        { ...native, results: undefined },
      ]),
    ).toThrow();
    expect(() => requireVitestPartitions(expected, [core, native, native])).toThrow();
    expect(() =>
      requireVitestPartitions(expected, [
        { ...core, exitCode: 1 },
        { ...native, exitCode: 2 },
      ]),
    ).toThrow('core: 1, raster: 2');
  });

  it('rejects missing, unexpected, duplicate, empty and incomplete suites', () => {
    expect(() => requireSuites([file], [passed])).not.toThrow();
    expect(() => requireSuites([], [])).toThrow('No test files');
    expect(() => requireSuites([file], [])).toThrow('Missing test result');
    expect(() => requireSuites([file], [{ ...passed, file: 'other.test.ts' }])).toThrow(
      'Unexpected',
    );
    expect(() => requireSuites([file], [passed, passed])).toThrow('Duplicate');
    for (const change of [{ passed: 0 }, { failed: 1 }, { skipped: 1 }, { todo: 1 }]) {
      expect(() => requireSuites([file], [{ ...passed, ...change }])).toThrow('Incomplete');
    }
    expect(() => requireSuites([file], [{ ...passed, passed: Number.NaN }])).toThrow(
      'Invalid count',
    );
  });

  it('discovers real TypeScript and inherited MJS suites without including helpers', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'office-kit-ci-'));
    try {
      await mkdir(join(directory, 'nested'));
      await Promise.all([
        writeFile(join(directory, 'a.test.ts'), ''),
        writeFile(join(directory, 'b.test.mjs'), ''),
        writeFile(join(directory, 'helper.ts'), ''),
        writeFile(join(directory, 'nested/c.test.ts'), ''),
      ]);
      expect(await testFiles(directory)).toEqual([
        join(directory, 'a.test.ts'),
        join(directory, 'b.test.mjs'),
      ]);
      expect(await testFiles(directory, true)).toHaveLength(3);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('runs the aggregate check and rejects failed, skipped, cancelled or missing jobs', async () => {
    const jobs = ['static', 'test', 'browser', 'fidelity', 'ooxml-oracle'];
    const results = Object.fromEntries(jobs.map((job) => [job, { result: 'success' }]));
    const runCheck = async (value: unknown): Promise<void> => {
      await execute(process.execPath, ['scripts/ci/required-check.ts'], {
        env: { ...process.env, CI_JOB_RESULTS: JSON.stringify(value) },
      });
    };
    await expect(runCheck(results)).resolves.toBeUndefined();
    for (const result of ['failure', 'skipped', 'cancelled']) {
      await expect(runCheck({ ...results, browser: { result } })).rejects.toThrow();
    }
    const { test: omitted, ...missing } = results;
    expect(omitted).toBeDefined();
    await expect(runCheck(missing)).rejects.toThrow();
    await expect(runCheck(null)).rejects.toThrow();
  });

  it('refuses partial unit runs and changing the required eight browser shards', async () => {
    await expect(
      execute(process.execPath, ['scripts/ci/node-tests.ts', 'editor-unit', '1/8']),
    ).rejects.toThrow();
    await expect(
      execute(process.execPath, ['scripts/ci/node-tests.ts', 'dev-browser', '1/2']),
    ).rejects.toThrow();
    await expect(
      execute(process.execPath, ['scripts/ci/node-tests.ts', 'dev-browser', '9/8']),
    ).rejects.toThrow();
  });
});
