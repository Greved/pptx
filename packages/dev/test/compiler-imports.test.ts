import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { pathToFileURL } from 'node:url';
import { createDeckCompiler } from '../src/bundle.ts';

test('compiled external imports load from file URLs with spaces and Unicode paths', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'office-compiler-'));
  const project = join(directory, 'spaced 日本語');
  await mkdir(project);
  const entry = join(project, 'deck.ts');
  const output = join(project, 'compiled.mjs');
  await writeFile(
    entry,
    "import { createPresentation } from '@office-kit/pptx'; export default createPresentation();",
  );
  const compiler = await createDeckCompiler(entry);
  try {
    await compiler.build(output);
    assert.match(await readFile(output, 'utf8'), /from ["']file:\/\//);
    const module: unknown = await import(pathToFileURL(output).href);
    assert.ok(typeof module === 'object' && module !== null && 'default' in module);
    assert.ok(typeof module.default === 'object' && module.default !== null);
  } finally {
    await compiler.close();
    await rm(directory, { recursive: true, force: true });
  }
});
