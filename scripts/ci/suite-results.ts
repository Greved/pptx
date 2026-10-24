import { readdir } from 'node:fs/promises';
import { resolve, relative } from 'node:path';

export interface SuiteResult {
  file: string;
  passed: number;
  failed: number;
  skipped: number;
  todo: number;
}

export async function testFiles(directory: string, recursive = false): Promise<string[]> {
  const files: string[] = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name);
    if (entry.isDirectory() && recursive) files.push(...(await testFiles(path, true)));
    else if (entry.isFile() && /\.test\.(?:ts|mjs)$/.test(entry.name)) files.push(path);
  }
  return files.sort();
}

/** Compare execution evidence with independently enumerated source files. */
export function requireSuites(expected: readonly string[], results: readonly SuiteResult[]): void {
  if (expected.length === 0) throw new Error('No test files discovered');
  const expectedSet = new Set(expected.map((file) => resolve(file)));
  const seen = new Set<string>();
  for (const result of results) {
    const file = resolve(result.file);
    if (!expectedSet.has(file)) throw new Error(`Unexpected test file: ${file}`);
    if (seen.has(file)) throw new Error(`Duplicate result: ${file}`);
    seen.add(file);
    for (const count of [result.passed, result.failed, result.skipped, result.todo]) {
      if (!Number.isSafeInteger(count) || count < 0) throw new Error(`Invalid count: ${file}`);
    }
    if (result.passed === 0 || result.failed > 0 || result.skipped > 0 || result.todo > 0) {
      throw new Error(
        `Incomplete suite ${relative(process.cwd(), file)}: ${JSON.stringify(result)}`,
      );
    }
  }
  for (const file of expectedSet) {
    if (!seen.has(file)) throw new Error(`Missing test result: ${file}`);
  }
}
