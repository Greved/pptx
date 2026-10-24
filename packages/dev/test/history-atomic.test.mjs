import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chmod, mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createHistory } from '../src/history.ts';

async function project(t) {
  const dir = await mkdtemp(join(tmpdir(), 'office-history-atomic-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  return dir;
}

test('an undo that cannot write every file leaves all files at their newer contents', async (t) => {
  const dir = await project(t);
  const deck = join(dir, 'deck.tsx');
  const locked = join(dir, 'locked');
  const data = join(locked, 'data.json');
  await mkdir(locked);
  await writeFile(deck, 'v1');
  await writeFile(data, 'v1');
  let writes = 0;
  const history = await createHistory(dir, () => {}, {
    writeFile: async (...args) => {
      writes += 1;
      if (writes === 2) throw Object.assign(new Error('disk write failed'), { code: 'EIO' });
      return writeFile(...args);
    },
  });
  await history.begin('agent', 'AI edit');
  await writeFile(deck, 'v2');
  await writeFile(data, 'v2');
  await history.end('agent');

  // Fail after a completed replacement so this exercises actual disk rollback,
  // including on Windows and when directory permissions cannot deny root.
  await assert.rejects(history.move('undo'), { code: 'EIO' });
  assert.equal(writes, 3);
  assert.equal(await readFile(deck, 'utf8'), 'v2');
  assert.equal(await readFile(data, 'utf8'), 'v2');
  for (const folder of [dir, locked])
    assert.deepEqual(
      (await readdir(folder)).filter((name) => name.endsWith('.tmp')),
      [],
    );
  assert.equal(history.state().undo, 'AI edit');
  assert.equal(history.state().redo, null);
  assert.equal(history.state().busy, false);

  await history.move('undo');
  assert.equal(await readFile(deck, 'utf8'), 'v1');
  assert.equal(await readFile(data, 'utf8'), 'v1');
});

test('undo and redo keep the permissions of the files they replace', async (t) => {
  const dir = await project(t);
  const script = join(dir, 'tool.mjs');
  await writeFile(script, 'v1');
  await chmod(script, 0o755);
  // Windows exposes the writable bit rather than POSIX executable permissions.
  const originalMode = (await stat(script)).mode & 0o777;
  const history = await createHistory(dir);
  await history.begin('agent', 'AI edit');
  await writeFile(script, 'v2');
  await history.end('agent');
  await history.move('undo');
  assert.equal(await readFile(script, 'utf8'), 'v1');
  assert.equal((await stat(script)).mode & 0o777, originalMode);
  await history.move('redo');
  assert.equal((await stat(script)).mode & 0o777, originalMode);
});
