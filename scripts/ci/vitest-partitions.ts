import { requireSuites, type SuiteResult } from './suite-results.ts';

export interface VitestPartition {
  readonly name: string;
  readonly files: readonly string[];
  readonly results: readonly SuiteResult[] | undefined;
  readonly exitCode: number | null;
}

export function requireVitestPartitions(
  expected: readonly string[],
  partitions: readonly VitestPartition[],
): SuiteResult[] {
  const failed = partitions.filter((partition) => partition.exitCode !== 0);
  if (failed.length)
    throw new Error(
      `Vitest partitions failed: ${failed.map((partition) => `${partition.name}: ${partition.exitCode}`).join(', ')}`,
    );
  const names = new Set<string>();
  const results: SuiteResult[] = [];
  for (const partition of partitions) {
    if (names.has(partition.name)) throw new Error(`Duplicate partition: ${partition.name}`);
    names.add(partition.name);
    if (partition.results === undefined)
      throw new Error(`Missing execution report for partition: ${partition.name}`);
    requireSuites(partition.files, partition.results);
    results.push(...partition.results);
  }
  requireSuites(expected, results);
  return results;
}
